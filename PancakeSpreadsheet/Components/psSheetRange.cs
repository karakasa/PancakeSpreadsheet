using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using PancakeSpreadsheet.Params;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public class psSheetRange : PancakeComponent
    {
        public override Guid ComponentGuid => new("{B19AC8ED-BB25-448A-997A-F4494D03AB95}");

        protected override string ComponentName => "Sheet Range";

        protected override string ComponentDescription => "Get the mimimum and maximum index of columns in a row or rows inside a sheet.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Row Index to Query", "R", "The row to query.\r\n-1 for rows in the sheet.", GH_ParamAccess.item, -1);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddIntegerParameter("First Index", "F", "First existing index of the queryed", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Last Index", "L", "Last existing index of the queryed", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooSheet gooSheet = default;
            int rowId = 0;

            DA.GetData(0, ref gooSheet);
            DA.GetData(1, ref rowId);

            var sheet = gooSheet?.Value;

            if (sheet is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid sheet.");
                return;
            }

            if (rowId < 0)
            {
                DA.SetData(0, sheet.FirstRowNum);
                DA.SetData(1, sheet.LastRowNum);
            }
            else
            {
                var row = sheet.GetRow(rowId);
                if (row is null)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Row doesn't exist.");
                    return;
                }

                DA.SetData(0, row.FirstCellNum);
                DA.SetData(1, row.LastCellNum - 1);
            }
        }
        protected override string ComponentCategory => PancakeComponent.CategorySheet;
        protected override Bitmap Icon => ComponentIcons.SheetRange;
        public override GH_Exposure Exposure => GH_Exposure.tertiary;
    }
}
