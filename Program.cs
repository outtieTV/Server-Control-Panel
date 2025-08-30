using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Security.Principal;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Forms;
using MaterialSkin;
using MaterialSkin.Controls;

namespace ServerControlPanel
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
    }

    public class MainForm : MaterialForm
    {
        // Data models
        public class ExecutableItem
        {
            public string Nickname { get; set; } = "";
            public string Path { get; set; } = "";
            public string Parameters { get; set; } = "";
            public bool StartOnLogin { get; set; } = false;
        }

        public class ServiceItem
        {
            public string Nickname { get; set; } = "";
            public string ServiceName { get; set; } = "";
            public bool StartOnLogin { get; set; } = false;
        }

        public class Profile
        {
            public string Name { get; set; } = "";
            public bool StartOnLogin { get; set; } = false;
            public List<ExecutableItem> Executables { get; set; } = new();
            public List<ServiceItem> Services { get; set; } = new();
        }

        // Storage
        private readonly string DataFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ServerControlPanel", "profiles.json");

        private List<Profile> Profiles = new();

        // UI: Tabs
        private readonly MaterialTabControl tabs = new() { Dock = DockStyle.Fill };
        private readonly TabPage tabProfiles = new() { Text = "Profiles" };
        private readonly TabPage tabEdit = new() { Text = "Edit Profiles" };
        private readonly MaterialTabSelector tabSelector;

        // UI: Tab 1 (Profiles)
        private readonly Panel leftPanel = new() { Dock = DockStyle.Left, Width = 250 };
        private readonly Panel middlePanel = new() { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.White };
        private readonly Panel actionPanel = new() { Dock = DockStyle.Right, Width = 180, AutoScroll = true, BackColor = Color.White };
        private readonly MaterialButton btnNewProfile = new() { Text = "NEW PROFILE", Dock = DockStyle.Top };
        private readonly ListBox lstProfiles = new() { Dock = DockStyle.Fill };

        private readonly Panel rightPanel = new() { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.White };

        // UI: Tab 2 (Edit Profiles)
        private readonly MaterialTextBox2 txtProfileName = new() { Hint = "Profile Name" };
        private readonly MaterialCheckbox chkStartOnLogin = new() { Text = "Start Application + Profile on Login" };

        // Executable editor
        private readonly MaterialTextBox2 txtExeNick = new() { Hint = "Exe Nickname" };
        private readonly MaterialTextBox2 txtExePath = new() { Hint = "Exe Path (Browse)", ReadOnly = true };
        private readonly MaterialTextBox2 txtExeParams = new() { Hint = "Exe Parameters (optional)" };
        private readonly MaterialButton btnBrowseExe = new() { Text = "Browse EXE" };
        private readonly MaterialButton btnAddExe = new() { Text = "Add/Update EXE" };
        private readonly MaterialCheckbox chkExeLogin = new() { Text = "Start on Login", Width=200};
        private readonly ListBox lstExes = new() { Height = 140 };
        private readonly MaterialButton btnRemoveExe = new() { Text = "Remove EXE" };

        // Service editor
        private readonly MaterialTextBox2 txtSvcNick = new() { Hint = "Service Nickname" };
        private readonly ComboBox cboServices = new() { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly MaterialButton btnAddSvc = new() { Text = "Add/Update Service" };
        private readonly MaterialCheckbox chkSvcLogin = new() { Text = "Start on Login", Width=200 };
        private readonly ListBox lstSvcs = new() { Height = 140 };
        private readonly MaterialButton btnRemoveSvc = new() { Text = "Remove Service" };

        // Actions
        private readonly MaterialButton btnSaveProfile = new() { Text = "Save Profile" };
        private readonly MaterialButton btnDeleteProfile = new() { Text = "Delete Profile" };

        // State for editing
        private Profile? editingProfile = null;

        public MainForm()
        {
            Text = "Server Control Panel";
            Width = 1100;
            Height = 720;
            StartPosition = FormStartPosition.CenterScreen;

            // MaterialSkin Theme (Light)
            var skin = MaterialSkinManager.Instance;
            skin.AddFormToManage(this);
            skin.Theme = MaterialSkinManager.Themes.LIGHT;
            skin.ColorScheme = new ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE);

            // Build UI
            tabSelector = new MaterialTabSelector { BaseTabControl = tabs, Dock = DockStyle.Top };
            Controls.Add(tabs);
            Controls.Add(tabSelector);

            tabs.Controls.Add(tabProfiles);
            tabs.Controls.Add(tabEdit);

            BuildProfilesTab();
            BuildEditTab();
            LoadProfiles();
            RenderProfileList();

            foreach (var p in Profiles.Where(p => p.StartOnLogin))
            {
                foreach (var ex in p.Executables.Where(e => e.StartOnLogin))
                    StartExecutable(ex);

                foreach (var sv in p.Services.Where(sv => sv.StartOnLogin))
                    StartService(sv.ServiceName);
            }

        }

        private bool IsRunningAsAdministrator()
        {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        private void BuildProfilesTab()
        {
            // Left panel (profiles list)
            leftPanel.Controls.Add(lstProfiles);
            leftPanel.Controls.Add(new Panel { Height = 8, Dock = DockStyle.Top });
            leftPanel.Controls.Add(btnNewProfile);
            lstProfiles.Font = new Font("Segoe UI", 10);

            // Add three panels
            tabProfiles.Controls.Add(middlePanel);
            tabProfiles.Controls.Add(actionPanel);
            tabProfiles.Controls.Add(leftPanel);

            btnNewProfile.Click += (_, __) =>
            {
                editingProfile = new Profile();
                FillEditFields(editingProfile);
                tabs.SelectedTab = tabEdit;
                txtProfileName.Focus();
            };

            lstProfiles.SelectedIndexChanged += (_, __) =>
            {
                if (lstProfiles.SelectedItem is Profile p)
                {
                    RenderMiddlePanel(p);
                    RenderActionPanel(p);
                    editingProfile = p;
                    FillEditFields(p);
                }
            };
        }

        private void BuildEditTab()
        {
            var host = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.White };
            tabEdit.Controls.Add(host);

            // Layout helpers
            int margin = 16;
            int x = margin, y = margin, w = 840, h = 56;

            void Place(Control c, int height = -1, int width = -1)
            {
                c.Left = x;
                c.Top = y;
                c.Width = width > 0 ? width : w;
                c.Height = height > 0 ? height : h;
                host.Controls.Add(c);
                y += c.Height + 8;
            }


            // Profile Name
            txtProfileName.ForeColor = Color.Black;
            Place(txtProfileName);
            Place(chkStartOnLogin, 36);

            // Executable editor group
            var lblExe = new MaterialLabel { Text = "Executables", ForeColor = Color.Black };
            Place(lblExe, 24, 200);

            Place(txtExeNick);
            Place(txtExeParams);
            Place(txtExePath);

            var exeRow = new FlowLayoutPanel { Left = x, Top = y, Width = w, Height = 60, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            exeRow.Controls.Add(btnBrowseExe);
            exeRow.Controls.Add(btnAddExe);
            exeRow.Controls.Add(chkExeLogin);
            exeRow.AutoScroll = true;
            host.Controls.Add(exeRow);
            y += exeRow.Height + 8;

            lstExes.Width = w;
            Place(lstExes, 140);
            Place(btnRemoveExe, 40, 200);

            // Service editor group
            y += 8;
            var lblSvc = new MaterialLabel { Text = "Services", ForeColor = Color.Black };
            Place(lblSvc, 24, 200);

            Place(txtSvcNick);

            var svcRow = new FlowLayoutPanel { Left = x, Top = y, Width = w, Height = 60, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            svcRow.Controls.Add(new MaterialLabel { Text = "Installed Services:", ForeColor = Color.Black, Width = 180, Height = 36, Margin = new Padding(0, 6, 8, 0) });
            cboServices.Width = 500;
            svcRow.Controls.Add(cboServices);
            svcRow.Controls.Add(btnAddSvc);
            svcRow.Controls.Add(chkSvcLogin);
            svcRow.AutoScroll = true;
            host.Controls.Add(svcRow);
            y += svcRow.Height + 8;

            lstSvcs.Width = w;
            Place(lstSvcs, 140);
            Place(btnRemoveSvc, 40, 200);

            // Save / Delete
            var actionRow = new FlowLayoutPanel { Left = x, Top = y, Width = w, Height = 48, FlowDirection = FlowDirection.LeftToRight, WrapContents = false };
            actionRow.Controls.Add(btnSaveProfile);
            actionRow.Controls.Add(btnDeleteProfile);
            host.Controls.Add(actionRow);

            // Events
            btnBrowseExe.Click += (_, __) =>
            {
                using var ofd = new OpenFileDialog
                {
                    Filter = "Executable (*.exe)|*.exe|All files (*.*)|*.*",
                    Title = "Select Executable"
                };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtExePath.Text = ofd.FileName;
                    if (string.IsNullOrWhiteSpace(txtExeNick.Text))
                        txtExeNick.Text = Path.GetFileNameWithoutExtension(ofd.FileName);
                }
            };

            btnAddExe.Click += (_, __) =>
            {
                if (editingProfile == null) editingProfile = new Profile();
                if (string.IsNullOrWhiteSpace(txtExeNick.Text) || string.IsNullOrWhiteSpace(txtExePath.Text))
                {
                    MessageBox.Show("Provide an EXE nickname and select an executable path.", "Missing info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                var item = new ExecutableItem
                {
                    Nickname = txtExeNick.Text.Trim(),
                    Path = txtExePath.Text.Trim(),
                    Parameters = txtExeParams.Text?.Trim() ?? "",
                    StartOnLogin = chkExeLogin.Checked
                };
                // Upsert by nickname
                var existing = editingProfile.Executables.FirstOrDefault(e => e.Nickname.Equals(item.Nickname, StringComparison.OrdinalIgnoreCase));
                if (existing != null)
                {
                    existing.Path = item.Path;
                    existing.Parameters = item.Parameters;
                }
                else
                {
                    editingProfile.Executables.Add(item);
                }
                RefreshEditLists();
                txtExeNick.Text = "";
                txtExePath.Text = "";
                txtExeParams.Text = "";
            };

            btnRemoveExe.Click += (_, __) =>
            {
                if (editingProfile == null) return;
                if (lstExes.SelectedItem is ExecutableItem ex)
                {
                    editingProfile.Executables.Remove(ex);
                    RefreshEditLists();
                }
            };

            btnAddSvc.Click += (_, __) =>
            {
                if (editingProfile == null) editingProfile = new Profile();
                if (string.IsNullOrWhiteSpace(txtSvcNick.Text) || cboServices.SelectedItem == null)
                {
                    MessageBox.Show("Provide a service nickname and select a service.", "Missing info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string svcName = (cboServices.SelectedItem as ServiceController)?.ServiceName
                                 ?? cboServices.SelectedItem.ToString()!;
                var item = new ServiceItem
                {
                    Nickname = txtSvcNick.Text.Trim(),
                    ServiceName = svcName,
                    StartOnLogin = chkSvcLogin.Checked
                };
                var existing = editingProfile.Services.FirstOrDefault(s => s.Nickname.Equals(item.Nickname, StringComparison.OrdinalIgnoreCase));
                if (existing != null) existing.ServiceName = item.ServiceName;
                else editingProfile.Services.Add(item);

                RefreshEditLists();
                txtSvcNick.Text = "";
                cboServices.SelectedIndex = -1;
            };

            btnRemoveSvc.Click += (_, __) =>
            {
                if (editingProfile == null) return;
                if (lstSvcs.SelectedItem is ServiceItem si)
                {
                    editingProfile.Services.Remove(si);
                    RefreshEditLists();
                }
            };

            btnSaveProfile.Click += (_, __) =>
            {
                if (editingProfile == null) editingProfile = new Profile();
                if (string.IsNullOrWhiteSpace(txtProfileName.Text))
                {
                    MessageBox.Show("Profile needs a name.", "Missing info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                editingProfile.Name = txtProfileName.Text.Trim();
                editingProfile.StartOnLogin = chkStartOnLogin.Checked;
                // Upsert into Profiles
                var existing = Profiles.FirstOrDefault(p => p.Name.Equals(editingProfile.Name, StringComparison.OrdinalIgnoreCase));
                if (existing != null)
                {
                    existing.Executables = editingProfile.Executables.ToList();
                    existing.Services = editingProfile.Services.ToList();
                }
                else
                {
                    Profiles.Add(editingProfile);
                }
                SaveProfiles();
                RegisterAppOnLogin(editingProfile.StartOnLogin);
                RenderProfileList();
                tabs.SelectedTab = tabProfiles;
            };

            btnDeleteProfile.Click += (_, __) =>
            {
                if (editingProfile == null || string.IsNullOrWhiteSpace(editingProfile.Name)) return;
                var res = MessageBox.Show($"Delete profile '{editingProfile.Name}'?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (res == DialogResult.Yes)
                {
                    Profiles.RemoveAll(p => p.Name.Equals(editingProfile!.Name, StringComparison.OrdinalIgnoreCase));
                    SaveProfiles();
                    RenderProfileList();
                    editingProfile = new Profile();
                    FillEditFields(editingProfile);
                }
            };

            // Populate installed services
            try
            {
                var services = ServiceController.GetServices().OrderBy(s => s.DisplayName).ToList();
                cboServices.Items.Clear();
                foreach (var s in services) cboServices.Items.Add(s);
                cboServices.DisplayMember = "DisplayName";
                cboServices.ValueMember = "ServiceName";
            }
            catch { /* Non-admin or restricted environment */ }

            // Listbox display members
            lstExes.DisplayMember = "Nickname";
            lstSvcs.DisplayMember = "Nickname";

            // Allow editing by clicking on an exe nickname
            lstExes.SelectedIndexChanged += (_, __) =>
            {
                if (lstExes.SelectedItem is ExecutableItem ex)
                {
                    txtExeNick.Text = ex.Nickname;
                    txtExePath.Text = ex.Path;
                    txtExeParams.Text = ex.Parameters;
                    chkExeLogin.Checked = ex.StartOnLogin;
                }
            };
            // Allow editing by clicking on a service nickname
            lstSvcs.SelectedIndexChanged += (_, __) =>
            {
                if (lstSvcs.SelectedItem is ServiceItem sv)
                {
                    txtSvcNick.Text = sv.Nickname;
                    // Find and select the service in the ComboBox
                    var service = cboServices.Items.Cast<ServiceController>().FirstOrDefault(s => s.ServiceName == sv.ServiceName);
                    if (service != null)
                    {
                        cboServices.SelectedItem = service;
                    }
                    chkSvcLogin.Checked = sv.StartOnLogin;
                }
            };
        }

        private void RenderProfileList()
        {
            lstProfiles.Items.Clear();
            foreach (var p in Profiles.OrderBy(p => p.Name))
                lstProfiles.Items.Add(p);
            lstProfiles.DisplayMember = "Name";
            if (lstProfiles.Items.Count > 0 && lstProfiles.SelectedIndex == -1)
                lstProfiles.SelectedIndex = 0;
        }

        private void RenderRightPanelForProfile(Profile p)
        {
            rightPanel.Controls.Clear();

            int margin = 16;
            int y = margin;

            // Title
            var title = new MaterialLabel
            {
                Text = $"Profile: {p.Name}",
                ForeColor = Color.Black,
                FontType = MaterialSkinManager.fontType.H5,
                Left = margin,
                Top = y
            };
            rightPanel.Controls.Add(title);
            y += 36;

            // Executables
            var exeHeader = new MaterialLabel { Text = "Executables", ForeColor = Color.Black, Left = margin, Top = y };
            rightPanel.Controls.Add(exeHeader);
            y += 28;

            foreach (var ex in p.Executables)
            {
                var row = CreateControlRowForExecutable(ex);
                row.Top = y;
                row.Left = margin;
                rightPanel.Controls.Add(row);
                y += row.Height + 10;
            }

            // Services
            y += 8;
            var svcHeader = new MaterialLabel { Text = "Services", ForeColor = Color.Black, Left = margin, Top = y };
            rightPanel.Controls.Add(svcHeader);
            y += 28;

            foreach (var sv in p.Services)
            {
                var row = CreateControlRowForService(sv);
                row.Top = y;
                row.Left = margin;
                rightPanel.Controls.Add(row);
                y += row.Height + 10;
            }
        }

        private Panel CreateControlRowForExecutable(ExecutableItem ex)
        {
            var panel = new Panel { Width = rightPanel.ClientSize.Width - 32, Height = 44, Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top };

            var lbl = new MaterialLabel
            {
                Text = ex.Nickname,
                ForeColor = Color.Black,
                Left = 0,
                Top = 10,
                AutoSize = true
            };

            var btnStart = new MaterialButton { Text = "START", Left = 260, Top = 4 };
            var btnStop = new MaterialButton { Text = "STOP", Left = 340, Top = 4 };

            btnStart.Click += (_, __) =>
            {
                StartExecutable(ex);
                lbl.ForeColor = Color.Green;
            };

            btnStop.Click += (_, __) =>
            {
                StopExecutable(ex);
                lbl.ForeColor = Color.Red;
            };

            // path preview
            var small = new Label
            {
                Text = $"{ex.Path} {ex.Parameters}".Trim(),
                ForeColor = Color.DimGray,
                AutoEllipsis = true,
                Left = 0,
                Top = 26,
                Width = panel.Width - 10,
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
            };

            panel.Controls.Add(lbl);
            panel.Controls.Add(btnStart);
            panel.Controls.Add(btnStop);
            panel.Controls.Add(small);
            return panel;
        }

        private Panel CreateControlRowForService(ServiceItem sv)
        {
            var panel = new Panel { Width = rightPanel.ClientSize.Width - 32, Height = 44, Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top };

            var lbl = new MaterialLabel
            {
                Text = sv.Nickname,
                ForeColor = Color.Black,
                Left = 0,
                Top = 10,
                AutoSize = true
            };

            var btnStart = new MaterialButton { Text = "START", Left = 260, Top = 4 };
            var btnStop = new MaterialButton { Text = "STOP", Left = 340, Top = 4 };

            btnStart.Click += (_, __) =>
            {
                StartService(sv.ServiceName);
                lbl.ForeColor = Color.Green;
            };

            btnStop.Click += (_, __) =>
            {
                StopService(sv.ServiceName);
                lbl.ForeColor = Color.Red;
            };

            var small = new Label
            {
                Text = sv.ServiceName,
                ForeColor = Color.DimGray,
                AutoEllipsis = true,
                Left = 0,
                Top = 26,
                Width = panel.Width - 10,
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
            };

            panel.Controls.Add(lbl);
            panel.Controls.Add(btnStart);
            panel.Controls.Add(btnStop);
            panel.Controls.Add(small);
            return panel;
        }

        // Process helpers
        private void StartExecutable(ExecutableItem ex)
        {
            try
            {
                if (!File.Exists(ex.Path))
                {
                    MessageBox.Show($"File not found:\n{ex.Path}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                // Use batch 'start' via cmd to mimic detached start with params
                string args = $"/c start \"\" \"{ex.Path}\" {ex.Parameters}".TrimEnd();
                var psi = new ProcessStartInfo("cmd.exe", args)
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    WorkingDirectory = Path.GetDirectoryName(ex.Path) ?? Environment.CurrentDirectory
                };
                Process.Start(psi);
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message, "Start EXE failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void StopExecutable(ExecutableItem ex)
        {
            try
            {
                var name = Path.GetFileNameWithoutExtension(ex.Path);
                var procs = Process.GetProcessesByName(name);
                if (procs.Length == 0)
                {
                    MessageBox.Show($"{name}.exe is not running.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                foreach (var p in procs)
                {
                    try { p.Kill(true); } catch { /* ignore */ }
                    try { p.WaitForExit(3000); } catch { /* ignore */ }
                }
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message, "Stop EXE failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void StartService(string serviceName)
        {
            try
            {
                using (ServiceController sc = new ServiceController(serviceName))
                {
                    if (sc.Status == ServiceControllerStatus.Stopped || sc.Status == ServiceControllerStatus.Paused)
                    {
                        sc.Start();
                        sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Please run this application as Administrator.", "Start Service failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void StopService(string serviceName)
        {
            try
            {
                using (ServiceController sc = new ServiceController(serviceName))
                {
                    if (sc.Status == ServiceControllerStatus.Running)
                    {
                        sc.Stop();
                        sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(30));
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Please run this application as Administrator.", "Stop Service failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // Persistence
        private void RegisterAppOnLogin(bool enable)
        {
            try
            {
                string runKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(runKey, true)!;
                if (enable)
                    key.SetValue("ServerControlPanel", Application.ExecutablePath);
                else
                    key.DeleteValue("ServerControlPanel", false);
            }
            catch { /* ignore */ }
        }
        private void LoadProfiles()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(DataFile)!);
                if (File.Exists(DataFile))
                {
                    var json = File.ReadAllText(DataFile);
                    var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                    Profiles = JsonSerializer.Deserialize<List<Profile>>(json, opts) ?? new List<Profile>();
                }
            }
            catch
            {
                Profiles = new List<Profile>();
            }
        }

        private void SaveProfiles()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(DataFile)!);
                var json = JsonSerializer.Serialize(Profiles, new JsonSerializerOptions { WriteIndented = true, DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull });
                File.WriteAllText(DataFile, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Save failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void RenderMiddlePanel(Profile p)
        {
            middlePanel.Controls.Clear();
            int margin = 16, y = margin;

            var title = new MaterialLabel
            {
                Text = $"Profile: {p.Name}",
                ForeColor = Color.Black,
                FontType = MaterialSkinManager.fontType.H5,
                Left = margin,
                Top = y
            };
            middlePanel.Controls.Add(title);
            y += 36;

            var exeHeader = new MaterialLabel { Text = "Executables", ForeColor = Color.Black, Left = margin, Top = y };
            middlePanel.Controls.Add(exeHeader);
            y += 28;

            foreach (var ex in p.Executables)
            {
                var lbl = new Label
                {
                    Text = $"{ex.Nickname}\n{ex.Path} {ex.Parameters}".Trim(),
                    ForeColor = Color.Black,
                    Left = margin,
                    Top = y,
                    AutoSize = false,
                    MaximumSize = new Size(middlePanel.Width - 32, 0), // word wrap
                    AutoEllipsis = false
                };
                lbl.Font = new Font("Segoe UI", 9);
                lbl.Width = middlePanel.Width - 32;
                lbl.Height = TextRenderer.MeasureText(lbl.Text, lbl.Font,
                    new Size(lbl.Width, int.MaxValue), TextFormatFlags.WordBreak).Height;

                middlePanel.Controls.Add(lbl);
                y += lbl.Height + 10;
            }

            y += 8;
            var svcHeader = new MaterialLabel { Text = "Services", ForeColor = Color.Black, Left = margin, Top = y };
            middlePanel.Controls.Add(svcHeader);
            y += 28;

            foreach (var sv in p.Services)
            {
                var lbl = new Label
                {
                    Text = $"{sv.Nickname}\n{sv.ServiceName}",
                    ForeColor = Color.Black,
                    Left = margin,
                    Top = y,
                    AutoSize = false,
                    MaximumSize = new Size(middlePanel.Width - 32, 0), // word wrap
                    AutoEllipsis = false
                };
                lbl.Font = new Font("Segoe UI", 9);
                lbl.Width = middlePanel.Width - 32;
                lbl.Height = TextRenderer.MeasureText(lbl.Text, lbl.Font,
                    new Size(lbl.Width, int.MaxValue), TextFormatFlags.WordBreak).Height;

                middlePanel.Controls.Add(lbl);
                y += lbl.Height + 10;
            }
        }
        private void RenderActionPanel(Profile p)
        {
            actionPanel.Controls.Clear();
            int margin = 16, y = margin;

            foreach (var ex in p.Executables)
            {
                var lbl = new MaterialLabel { Text = ex.Nickname, Left = margin, Top = y, ForeColor = Color.Black };
                var btnStart = new MaterialButton { Text = "START", Left = margin, Top = y + 24, Width = 70 };
                var btnStop = new MaterialButton { Text = "STOP", Left = margin + 80, Top = y + 24, Width = 70 };

                btnStart.Click += (_, __) => StartExecutable(ex);
                btnStop.Click += (_, __) => StopExecutable(ex);

                actionPanel.Controls.Add(lbl);
                actionPanel.Controls.Add(btnStart);
                actionPanel.Controls.Add(btnStop);
                y += 70;
            }

            foreach (var sv in p.Services)
            {
                var lbl = new MaterialLabel { Text = sv.Nickname, Left = margin, Top = y, ForeColor = Color.Black };
                var btnStart = new MaterialButton { Text = "START", Left = margin, Top = y + 24, Width = 70 };
                var btnStop = new MaterialButton { Text = "STOP", Left = margin + 80, Top = y + 24, Width = 70 };

                btnStart.Click += (_, __) => StartService(sv.ServiceName);
                btnStop.Click += (_, __) => StopService(sv.ServiceName);

                actionPanel.Controls.Add(lbl);
                actionPanel.Controls.Add(btnStart);
                actionPanel.Controls.Add(btnStop);
                y += 70;
            }
        }

        // Edit helpers
        private void FillEditFields(Profile p)
        {
            txtProfileName.Text = p.Name ?? "";
            editingProfile = new Profile
            {
                Name = p.Name ?? "",
                Executables = p.Executables.Select(e => new ExecutableItem
                {
                    Nickname = e.Nickname,
                    Path = e.Path,
                    Parameters = e.Parameters
                }).ToList(),
                Services = p.Services.Select(s => new ServiceItem
                {
                    Nickname = s.Nickname,
                    ServiceName = s.ServiceName
                }).ToList()
            };
            RefreshEditLists();
        }

        private void RefreshEditLists()
        {
            if (editingProfile == null) return;

            lstExes.Items.Clear();
            foreach (var e in editingProfile.Executables)
                lstExes.Items.Add(e);

            lstSvcs.Items.Clear();
            foreach (var s in editingProfile.Services)
                lstSvcs.Items.Add(s);

            lstExes.DisplayMember = "Nickname";
            lstSvcs.DisplayMember = "Nickname";
        }

        // ToString overrides for listboxes (fallback)
        public override string ToString() => base.ToString();
    }
}
