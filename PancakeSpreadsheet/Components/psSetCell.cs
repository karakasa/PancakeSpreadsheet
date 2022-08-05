using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
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
    public class psSetCell : PancakeComponent
    {
        public override Guid ComponentGuid => new("{B19AC8ED-AA25-448A-997A-F4494D03AB96}");

        protected override string ComponentName => "Set Cell";

        protected override string ComponentDescription => "Set the content of a specific cell.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
            pManager.AddParameter(new ParamCellReference(), "Cell", "C", "Cell reference", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Type", "T", "Cell type for content\r\nRight-click for more options.", GH_ParamAccess.item, 0);

            this.AddCellTypeHintValues();

            pManager.AddGenericParameter("Content", "C", "Content to set.", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooSheet gooSheet = default;
            GooCellReference gooReferences = default;
            int option = 0;
            IGH_Goo content = default;

            DA.GetData(0, ref gooSheet);
            DA.GetData(1, ref gooReferences);
            DA.GetData(2, ref option);
            DA.GetData(3, ref content);

            var sheet = gooSheet?.Value;

            if (sheet is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid sheet.");
                return;
            }

            if (gooReferences is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid cell reference.");
                return;
            }

            var reference = gooReferences.Value;

            var row = sheet.GetRow(reference.RowId);

            if (row is null)
                row = sheet.CreateRow(reference.RowId);

            var cell = row.GetCell(reference.ColumnId);

            if (cell is null)
                cell = row.CreateCell(reference.ColumnId);

            try
            {
                var hint = CellAccessUtility.GetHint(option);
                if (!CellAccessUtility.TrySetCellContent(cell, hint, content))
                {
                    if (hint == CellTypeHint.Automatic)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Input is of a unknown type. It has been processed as text.");
                    }
                    else if (hint == CellTypeHint.Datetime)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Input cannot be converted to DateTime.");
                    }
                }

                DA.SetData(0, gooSheet);
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Something goes wrong.\r\n" + ex.ToString());
            }
        }
        protected override string ComponentCategory => PancakeComponent.CategoryCellContent;
        protected override Bitmap Icon => ComponentIcons.SetCell;
    }
}
