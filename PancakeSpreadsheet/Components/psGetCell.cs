using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
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
    public class psGetCell : PancakeComponent
    {
        protected override string ComponentCategory => PancakeComponent.CategoryCellContent;
        public override Guid ComponentGuid => new("{B19AC8ED-AA25-448A-997A-F4494D03AB95}");

        protected override string ComponentName => "Get Cell";

        protected override string ComponentDescription => "Get the content of a specific cell.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
            pManager.AddParameter(new ParamCellReference(), "Cell", "C", "Cell reference", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Type", "T", "Cell type for content retrieval\r\nRight-click for more options.", GH_ParamAccess.item, 0);

            this.AddCellTypeHintValues();
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Content", "C", "Cell content", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooSheet gooSheet = default;
            GooCellReference gooReferences = default;
            int option = 0;

            DA.GetData(0, ref gooSheet);
            DA.GetData(1, ref gooReferences);
            DA.GetData(2, ref option);

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

            var cell = sheet.GetRow(reference.RowId)?.GetCell(reference.ColumnId);

            if (cell is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "The cell doesn't exist.");

                DA.SetData(0, null);
                return;
            }

            object content = null;
            try
            {
                var hint = CellAccessUtility.GetHint(option);
                if (!CellAccessUtility.TryGetCellContent(cell, hint, out content))
                {
                    if (hint == CellTypeHint.Automatic)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"{reference} contains unknown type of data.");
                    }
                    else if (hint == CellTypeHint.Formula)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"{reference} is not a formula cell.");
                    }
                }
                else
                {
                    DA.SetData(0, content);
                }
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Something goes wrong.\r\n" + ex.ToString());
            }

            DA.SetData(0, content);
        }

        protected override Bitmap Icon => ComponentIcons.GetCell;
    }
}
