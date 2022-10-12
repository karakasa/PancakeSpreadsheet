using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
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
    public class psMoveCell : PancakeComponent
    {
        public override Guid ComponentGuid => new("{8110A674-1B1A-47F1-8B33-11C840AADE2A}");

        protected override string ComponentName => "Move Cell Area";

        protected override string ComponentDescription => "Move a cell or cell range reference according to row/column offsets.";

        protected override string ComponentCategory => PancakeComponent.CategoryCellContent;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Cell(s)", "C(R)", "Cell (range) to move.", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Row offset", "R", "Row offset. Can be negative.", GH_ParamAccess.item, 0);
            pManager.AddIntegerParameter("Column offset", "C", "Column offset. Can be negative.", GH_ParamAccess.item, 0);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Cell(s)", "C(R)", "Moved cell (range).", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            IGH_Goo goo = default;
            int rowOffset = 0;
            int colOffset = 0;
            DA.GetData(0, ref goo);
            DA.GetData(1, ref rowOffset);
            DA.GetData(2, ref colOffset);

            if (rowOffset == 0 && colOffset == 0)
            {
                DA.SetData(0, goo);
                return;
            }

            try
            {
                switch (goo)
                {
                    case GH_String ghStr:
                        var str = ghStr.Value;
                        if (SimpleCellReference.TryFromString(str, out var cell))
                        {
                            DA.SetData(0, Move(cell, rowOffset, colOffset).AsGoo());
                        }
                        else if (SimpleCellRange.TryFromString(str, out var cellRange))
                        {
                            DA.SetData(0, Move(cellRange, rowOffset, colOffset).AsGoo());
                        }
                        else
                        {
                            AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Cannot interpret the string input.");
                        }
                        break;
                    case GooCellReference gooCell:
                        DA.SetData(0, Move(gooCell.Value, rowOffset, colOffset).AsGoo());
                        break;
                    case GooCellRangeReference gooCellRange:
                        DA.SetData(0, Move(gooCellRange.Value, rowOffset, colOffset).AsGoo());
                        break;
                    default:
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Unrecognizable input.");
                        break;
                }
            }
            catch (NotSupportedException ex1)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, ex1.Message);
            }
            catch (IndexOutOfRangeException)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Given offsets make the cell reference invalid.");
            }
        }
        private static SimpleCellReference Move(SimpleCellReference cell, int row, int col)
        {
            cell.RowId += row;
            cell.ColumnId += col;

            if (!cell.IsValid())
                throw new IndexOutOfRangeException();

            return cell;
        }
        private static SimpleCellRange Move(SimpleCellRange cell, int row, int col)
        {
            cell.StartCell.RowId += row;
            cell.StartCell.ColumnId += col;

            cell.EndCell.RowId += row;
            cell.EndCell.ColumnId += col;

            if (!cell.IsValid())
                throw new IndexOutOfRangeException();

            return cell;
        }
        public override GH_Exposure Exposure => GH_Exposure.tertiary;
        protected override Bitmap Icon => ComponentIcons.MoveCells;
    }
}

