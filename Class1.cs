dotnet new sln --name unlockvba
dotnet new console --name project1
dotnet sln unlockvba.sln add project1\project1.csproj

Edit .csproj to add dll

>dotnet build

add to csproj
<ItemGroup>
    <Reference Include="OpenMcdf">
      <SpecificVersion>False</SpecificVersion>
      <HintPath>..\OpenMcdf.dll</HintPath>
    </Reference>
</ItemGroup>
