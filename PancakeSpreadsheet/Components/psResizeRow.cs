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
    public class psResizeRow : PancakeComponent
    {
        public override GH_Exposure Exposure => GH_Exposure.tertiary;
        public override Guid ComponentGuid => new("{FCA4CEA6-12D7-40BB-A870-561B538D8717}");

        protected override string ComponentName => "Resize Row Height";

        protected override string ComponentDescription => "Resize row height.";

        protected override string ComponentCategory => CategoryCellStyle;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
            pManager.AddGenericParameter("Identifier", "I", "Identifier of the row(s) to be adjusted\r\nIt can be an index, or a cell reference, or a cell range.", GH_ParamAccess.item);
            pManager.AddNumberParameter("Height", "H", "Height of row\r\nSpecial values:\r\n0: Default height", GH_ParamAccess.item);
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
                var index = CellAccessUtility.GetRCIndex(gooIdentifier, true, out var isArray, out var array);

                if (isArray)
                {
                    foreach (var index2 in array)
                        SetRowHeight(sheet, index2, quantity);
                }
                else
                {
                    SetRowHeight(sheet, index, quantity);
                }

                DA.SetData(0, gooSheet);
            }
            catch (ArgumentException)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid identifier.");
                return;
            }
        }

        private static void SetRowHeight(ISheet sheet, int index, double quantity)
        {
            var row = sheet.GetRow(index) ?? sheet.CreateRow(index);

            if (Math.Abs(quantity) < 1e-7)
            {
                row.Height = sheet.DefaultRowHeight;
            }
            else if (Math.Abs(quantity + 1) < 1e-7)
            {
                row.Height = (short)-1;
            }
            else
            {
                row.HeightInPoints = (float)quantity;
            }
        }
        protected override Bitmap Icon => ComponentIcons.AdjustRowHeight;
    }
}
