using Grasshopper.Kernel;
using PancakeSpreadsheet.NpoiInterop;
using PancakeSpreadsheet.Params;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public class psSheetEntireRange : PancakeComponent
    {
        public override Guid ComponentGuid => new("{49B3E29B-FEA0-4593-B0DA-33DC4AA28D8A}");

        protected override string ComponentName => "Sheet Entire Range";

        protected override string ComponentDescription => "Get the entire range of a sheet.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddParameter(new ParamCellRangeReference(), "Cell Range", "CR", "Cell range reference", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooSheet gooSheet = default;

            DA.GetData(0, ref gooSheet);

            var sheet = gooSheet?.Value;

            if (sheet is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid sheet.");
                return;
            }

            var crange = Features.GetSheetRange(sheet);

            DA.SetData(0, crange.AsGoo());
        }
        protected override string ComponentCategory => PancakeComponent.CategorySheet;
        protected override Bitmap Icon => ComponentIcons.SheetEntireRange;
        public override GH_Exposure Exposure => GH_Exposure.tertiary;
    }
}
