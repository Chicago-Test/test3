============================================================
VSCode Portable
Create "Data" folder first.
Install .NET Core 3.x 
Install VSCode extensions. (C# by Microsoft/NuGet Package Manager /NuGet Gallery /)
Use Open folder to open .csproj/May need to edit "launch.json" for debugging/
Short-cut to compile (build) a C# app in Visual Studio Code (VSCode) is SHIFT+CTRL+B.
Install necessary extensions. Press 'F1' to do many things.
============================================================
Build a .net project without VisualStudio

>dotnet new sln --name yojitest
Don’t open the .sln… No
>dotnet new --help | more
>dotnet new console --name project1   (Console Application)
( >dotnet new classlib --name project1   (dll library)
To add a project to a solution file, you have to use dotnet sln command.
>dotnet sln yojitest.sln add project1\project1.csproj
>dotnet build
If the projects are referenced by the solution file -> dotnet restore
To add projects references on each projects -> dotnet add xxxx.csproj reference yyy.csproj

Change EXE to DLL-> Modify in csproj: "<OutputType>Exe</OutputType>" to "<OutputType>Library</OutputType>"
============================================================
To build multiple C# projects.
1. In the sln directory >dotnet build or,
2. Edit tasks.json
   a. Remove the specific path. "${workspaceFolder}/LMM/LMM.csproj" -> "${workspaceFolder}" or,
   b. Another way would be to have two build tasks.
============================================================
To change default startup program.
	In launch.json: "program": "${workspaceFolder}/mainTest/mainTest.exe",
To change .NET Core to .NET framework (Need .NET Framework SDK be installed)
	In csproj: <TargetFramework>net48</TargetFramework> (other options:"netcoreapp3.1")
============================================================
