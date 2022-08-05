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
    public class psGetMergedCells : PancakeComponent
    {
        public override Guid ComponentGuid => new("{0DC7F94E-2A72-4D03-938E-6359D3C7A1AD}");

        protected override string ComponentName => "Get Merged Cells";

        protected override string ComponentDescription => "Get merged cell ranges in a sheet.";

        protected override string ComponentCategory => PancakeComponent.CategoryCellStyle;

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddParameter(new ParamCellRangeReference(), "Cell Range", "CR", "Merged cell range", GH_ParamAccess.list);
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

            var regions = sheet.MergedRegions
                .Select(static region => SimpleCellRange.FromNpoiObj(region).AsGoo());

            DA.SetDataList(0, regions);
        }
        protected override Bitmap Icon => ComponentIcons.MergedCellList;
    }
}
