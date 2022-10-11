using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
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
    public class psCloneCellSytle : PancakeComponent
    {
        protected override string ComponentCategory => PancakeComponent.CategoryCellStyle;
        public override Guid ComponentGuid => new("{C8922373-306F-40C1-B54D-11A0FD3D49A7}");

        protected override string ComponentName => "Clone Cell Style";

        protected override string ComponentDescription => "Apply the style of a specific cell to one or more other cells.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S(T)", "Target sheet object", GH_ParamAccess.item);
            pManager.AddGenericParameter("Cell(s)", "C/CR(T)", "Target cell reference or cell range reference", GH_ParamAccess.item);

            pManager.AddGenericParameter("Source Sheet", "S(S)", "Source sheet object", GH_ParamAccess.item);
            pManager.AddParameter(new ParamCellReference(), "Source Cell", "C(S)", "Source cell reference", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooSheet gooSheet = default;
            GooSheet gooSheet2 = default;

            DA.GetData(0, ref gooSheet);

            var sheet = gooSheet?.Value;

            if (sheet is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid sheet.");
                return;
            }

            DA.GetData(2, ref gooSheet2);

            var srcSheet = gooSheet2?.Value;

            if (srcSheet is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid source sheet.");
                return;
            }

            if (!object.ReferenceEquals(srcSheet.Workbook, sheet.Workbook))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Source cell must be from the same spreadsheet.");
                return;
            }

            GooCellReference gooCellRef = default;

            DA.GetData(3, ref gooCellRef);

            var srcCref = gooCellRef.Value;

            IGH_Goo gooTarget = null;
            SimpleCellReference[] appliedCells;

            DA.GetData(1, ref gooTarget);

            appliedCells = CellAccessUtility.TryGetCellData(gooTarget);

            if (appliedCells is null || appliedCells.Length == 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Input must be a cell reference or a cell range reference");
                return;
            }

            var srcCell = srcSheet.GetRow(srcCref.RowId)?.GetCell(srcCref.ColumnId);

            if (srcCell is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Source cell {srcCell} doesn't exist.");
                return;
            }

            var style = srcCell.CellStyle;

            foreach (var cellRef in appliedCells)
            {
                var cell = sheet.EnsureCell(cellRef.RowId, cellRef.ColumnId);

                cell.CellStyle = style;
            }

            DA.SetData(0, gooSheet);
        }
        protected override Bitmap Icon => ComponentIcons.CopyCellStyle;
    }
}
