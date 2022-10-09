using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using PancakeSpreadsheet.NpoiInterop;
using PancakeSpreadsheet.Params;
using PancakeSpreadsheet.Utility;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public class psSetMultiCellsOld : PancakeComponent
    {
        public override Guid ComponentGuid => new("{AAEAC8ED-AA25-448A-997A-F4494D03AB95}");

        protected override string ComponentNickname => "psSetMultiCells";
        protected override string ComponentName => "Set Multiple Cells";

        protected override string ComponentDescription => "Set the content of a range of cells.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
            pManager.AddParameter(new ParamCellRangeReference(), "Cell Range", "CR", "Cell range reference", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Type", "T", "Cell type for content retrieval\r\nRight-click for more options.", GH_ParamAccess.item, 0);
            this.AddCellTypeHintValues();
            pManager.AddGenericParameter("Content", "C", "Content to set", GH_ParamAccess.tree);
            pManager.AddBooleanParameter("Row First", "R?", "True for row-first data organization; false for column-first.", GH_ParamAccess.item, true);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooSheet gooSheet = default;
            GooCellRangeReference gooReferences = default;
            int option = 0;
            bool rowFirst = true;

            DA.GetData(0, ref gooSheet);
            DA.GetData(1, ref gooReferences);
            DA.GetData(2, ref option);
            DA.GetDataTree<IGH_Goo>(3, out var data);
            DA.GetData(4, ref rowFirst);

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

            // TODO: Boundary Check
            var cntFirstLevel = crange.RowCount;
            var cntSecondLevel = crange.ColumnCount;

            if (!rowFirst)
                StaticExtensions.Swap(ref cntFirstLevel, ref cntSecondLevel);

            var dataCntFirstLevel = data.PathCount;
            var dataCntSecondLevel = data.Branches.Min(branch => branch.Count);

            if (dataCntFirstLevel != cntFirstLevel)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "First level counts of data mismatch. Data may be truncated.");
            }

            if(dataCntSecondLevel != cntSecondLevel)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Second level counts of data mismatch. Data may be truncated.");
            }

            var cellPositions = rowFirst ? crange.EnumerateRowFirst() : crange.EnumerateColumnFirst();

            var hint = CellAccessUtility.GetHint(option);

            using var dataEnumerator = data.Branches.GetEnumerator();

            foreach (var branch in cellPositions)
            {
                if (!dataEnumerator.MoveNext())
                    break;

                var curList = dataEnumerator.Current;
                var index = 0;

                foreach (var cref in branch)
                {
                    if (index >= curList.Count)
                        break;

                    var cell = sheet.EnsureCell(cref.RowId, cref.ColumnId);
                    if (!CellAccessUtility.TrySetCellContent(cell, hint, curList[index]))
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Cannot set the content of cell {cref}. Skipped.");
                    }

                    ++index;
                }
            }

            DA.SetData(0, gooSheet);
        }
        protected override string ComponentCategory => PancakeComponent.CategoryCellContent;
        protected override Bitmap Icon => ComponentIcons.SetCellRange;
        public override GH_Exposure Exposure => GH_Exposure.hidden;
        public override bool Obsolete => true;
    }
}
