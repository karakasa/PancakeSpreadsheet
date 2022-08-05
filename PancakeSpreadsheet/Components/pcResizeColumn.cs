using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using NPOI.SS.UserModel;
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
    public class psResizeColumn : PancakeComponent
    {
        public override GH_Exposure Exposure => GH_Exposure.tertiary;
        public override Guid ComponentGuid => new("{FCA4CEA6-121D-40BB-A870-561B538D8717}");

        protected override string ComponentName => "Resize Column Width";

        protected override string ComponentDescription => "Resize column width.";

        protected override string ComponentCategory => CategoryCellStyle;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
            pManager.AddGenericParameter("Identifier", "I", "Identifier of the column(s) to be adjusted\r\nIt can be an index, or a cell reference, or a cell range.", GH_ParamAccess.item);
            pManager.AddNumberParameter("Width", "W", "Width of column\r\nSpecial values:\r\n0: Default height; -1: Automatic sizing", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooSheet gooSheet = default;
            IGH_Goo gooIdentifier = null;
            double quantity = 0.0;

            DA.GetData(0, ref gooSheet);
            DA.GetData(1, ref gooIdentifier);
            DA.GetData(2, ref quantity);

            var sheet = gooSheet?.Value;

            if (sheet is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid sheet.");
                return;
            }

            try
            {
                var index = CellAccessUtility.GetRCIndex(gooIdentifier, false, out var isArray, out var array);

                if (isArray)
                {
                    foreach (var index2 in array)
                        SetColumnWidth(sheet, index2, quantity);
                }
                else
                {
                    SetColumnWidth(sheet, index, quantity);
                }

                DA.SetData(0, gooSheet);
            }
            catch (ArgumentException)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid identifier.");
                return;
            }
        }

        private static void SetColumnWidth(ISheet sheet, int index, double quantity)
        {
            // var row = sheet.GetRow(index) ?? sheet.CreateRow(index);

            if (Math.Abs(quantity) < 1e-7)
            {
                sheet.SetColumnWidth(index, sheet.DefaultColumnWidth);
            }
            else if (Math.Abs(quantity + 1) < 1e-7)
            {
                sheet.AutoSizeColumn(index);
            }
            else
            {
                sheet.SetColumnWidth(index, (int)Math.Round(quantity * 264.5195920558239));
            }
        }

        protected override Bitmap Icon => ComponentIcons.AdjustColumnWidth;
    }
}
