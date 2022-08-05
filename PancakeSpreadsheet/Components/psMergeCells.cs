using Grasshopper.Kernel;
using PancakeSpreadsheet.Params;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public class psMergeCells : PancakeComponent
    {
        public override Guid ComponentGuid => new("{0DC7F94E-2A72-4D03-938E-6359D3C7A1AC}");

        protected override string ComponentName => "Merge Cells";

        protected override string ComponentDescription => "Merge multiple cells into one.";

        protected override string ComponentCategory => PancakeComponent.CategoryCellStyle;

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
            pManager.AddParameter(new ParamCellRangeReference(), "Cell Range", "CR", "Cell range reference", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooSheet gooSheet = default;
            GooCellRangeReference gooReferences = default;

            DA.GetData(0, ref gooSheet);
            DA.GetData(1, ref gooReferences);

            var sheet = gooSheet?.Value;

            if (sheet is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid sheet.");
                return;
            }

            if (gooReferences is null || !gooReferences.Value.IsValid())
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid cell range reference.");
                return;
            }

            var crange = gooReferences.Value;

            if (crange.ColumnCount == 1 && crange.RowCount == 1)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Input must be at least 2 cells.");
                return;
            }

            try
            {
                sheet.AddMergedRegion(crange.AsNpoiObj());
                DA.SetData(0, gooSheet);
            }
            catch (InvalidOperationException invalidOperationEx)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Cannot merge cells.\r\n" + invalidOperationEx.Message);
            }
        }
        protected override Bitmap Icon => ComponentIcons.MergeCell;
    }
}
