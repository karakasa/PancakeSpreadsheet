using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace PancakeSpreadsheet.GH
{
    public class PluginInfo : GH_AssemblyInfo
    {
        public override string Name => "PancakeSpreadsheet";

        //Return a 24x24 pixel bitmap to represent this GHA library.
        public override Bitmap Icon => null;

        //Return a short string describing the purpose of this GHA library.
        public override string Description => "Provides dependency-less access to spreadsheet files.";

        public override Guid Id => new Guid("889F0F88-85A6-482E-8313-003A7416A256");

        //Return a string identifying you or your company.
        public override string AuthorName => "Keyu Gan";

        //Return a string representing your preferred contact details.
        public override string AuthorContact => "pancake@gankeyu.com";
    }
}