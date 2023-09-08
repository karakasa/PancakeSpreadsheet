using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

[assembly: GH_Loading(GH_LoadingDemand.ForceDirect)]

namespace EmbeddedEditor
{
    public class EmbeddedEditorInfo : GH_AssemblyInfo
    {
        public override string Name => "PancakeEmbeddedEditor";

        //Return a 24x24 pixel bitmap to represent this GHA library.
        public override Bitmap Icon => null;

        //Return a short string describing the purpose of this GHA library.
        public override string Description => "Provides embedded spreadsheet editor.";

        public override Guid Id => new Guid("e5fe0040-a3e2-4b36-a3e6-23d0810924cd");

        //Return a string identifying you or your company.
        public override string AuthorName => "Keyu Gan";

        //Return a string representing your preferred contact details.
        public override string AuthorContact => "pancake@gankeyu.com";
    }
}