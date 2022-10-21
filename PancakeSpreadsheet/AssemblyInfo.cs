using System.Resources;
using System.Reflection;
using System.Runtime.InteropServices;
using Grasshopper.Kernel;


// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("PancakeSpreadsheet")]
[assembly: AssemblyDescription("Import & export xlsx and xls files.")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("Keyu Gan")]
[assembly: AssemblyProduct("Pancake")]
[assembly: AssemblyCopyright("Copyright © Keyu Gan 2021")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("en-US")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
// [assembly: Guid("889f0f88-85a6-482e-8313-003a7416a25")] // This will also be the Guid of the Rhino plug-in

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion("1.1.0.0")]
[assembly: AssemblyFileVersion("1.1.0.0")]
[assembly: NeutralResourcesLanguage("en")]

[assembly: GH_Loading(GH_LoadingDemand.ForceDirect)]