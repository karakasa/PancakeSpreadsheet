﻿using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
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
    public class psGetMultiCells : PancakeComponent
    {
        protected override string ComponentCategory => PancakeComponent.CategoryCellContent;
        public override Guid ComponentGuid => new("{C19AC8ED-AA25-448A-997A-F4494D03AB95}");

        protected override string ComponentName => "Get Multiple Cells";

        protected override string ComponentDescription => "Get the content of a range of cells.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
            pManager.AddParameter(new ParamCellRangeReference(), "Cell Range", "CR", "Cell range reference", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Type", "T", "Cell type for content retrieval\r\nRight-click for more options.", GH_ParamAccess.item, 0);
            this.AddCellTypeHintValues();
            pManager.AddBooleanParameter("Row First", "R?", "True for row-first data organization; false for column-first.", GH_ParamAccess.item, true);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Content", "C", "Cell content", GH_ParamAccess.tree);
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
            DA.GetData(3, ref rowFirst);

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

            var cellPositions = rowFirst ? crange.EnumerateRowFirst() : crange.EnumerateColumnFirst();

            var hint = CellAccessUtility.GetHint(option);

            var tree = new GH_Structure<IGH_Goo>();
            var id = 0;
            foreach (var branch in cellPositions)
            {
                var list = new List<IGH_Goo>();

                foreach (var cref in branch)
                {
                    try
                    {
                        object content;

                        var cell = sheet.GetRow(cref.RowId)?.GetCell(cref.ColumnId);
                        if (cell is null)
                        {
                            content = null;
                        }
                        else
                        {
                            if (!CellAccessUtility.TryGetCellContent(cell, hint, out content))
                                content = null;
                        }

                        list.Add(GH_Convert.ToGoo(content));
                    }
                    catch (Exception ex)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Failed to access {cref}. Replaced with null.");
                        list.Add(null);
                    }
                }

                tree.EnsurePath(id).AddRange(list);
                ++id;
            }

            DA.SetDataTree(0, tree);
        }
        protected override Bitmap Icon => ComponentIcons.GetCellRange;
    }
}
