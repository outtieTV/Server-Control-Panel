# OuttieTV's Server Control Panel
A server control panel written in C# with support for many profiles.<br />
Features:<br />
- GUI for starting and stopping servers
- add as many profiles as you want
- add as many executables to a profile as you want
- add as many services to a profile as you want
- start on login for application, profiles, executables, and services
- edit profiles
- save profiles for next launch in %APPDATA% as a .json file
- theme by materialskin.2 https://www.nuget.org/packages/MaterialSkin.2
<br />
Instructions:<br />
1. install Visual Studio 2022 with C# packages<br />
2. open .sln file<br />
3. open nuget package manager and install materialskin.2<br />
```dotnet add package MaterialSkin.2 --version 2.3.1```<br />
4. build and run the program with RELEASE : ANY CPU<br />
<br />
Known issues:<br />
-executable full path sometimes doesn't show until you click around a bit. This is either due to RenderPanelLists or RenderMiddlePanel<br />
<img width="1100" height="720" alt="image" src="https://github.com/user-attachments/assets/a8cb4d3b-ea9a-418e-9838-20a6f7c7cf03" />
<img width="1920" height="1020" alt="image" src="https://github.com/user-attachments/assets/3dbc5f67-97b2-442d-b5e3-f329964dab33" />

