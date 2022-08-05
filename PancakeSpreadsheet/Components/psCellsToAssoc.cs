using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using PancakeSpreadsheet.NpoiInterop;
using PancakeSpreadsheet.PancakeInterop;
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
    public class psCellsToAssoc : PancakeComponentRequireCore
    {
        protected override string ComponentCategory => PancakeComponent.CategoryCellContent;
        public override Guid ComponentGuid => new("{C19AC8ED-BB01-448A-997A-F4494D03AB95}");

        protected override string ComponentName => "Cells to Assoc";

        protected override string ComponentDescription => "Get the content of a range of cells, as associations.\r\n" +
            "This component requires the core Pancake experience.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
            pManager.AddParameter(new ParamCellRangeReference(), "Cell Range", "CR", "Cell range reference", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Type", "T", "Cell type for content retrieval\r\nRight-click for more options.", GH_ParamAccess.item, 0);
            this.AddCellTypeHintValues();
            pManager.AddBooleanParameter("Row First", "R?", "True for row-first data organization; false for column-first.", GH_ParamAccess.item, true);

            pManager.AddBooleanParameter("Has Header", "H?", "Whether the data includes headers. True by default.\r\n" +
                "If false, you may designate alternative headers.", GH_ParamAccess.item, true);

            pManager.AddTextParameter("Alternative Headers", "AH", "Alternative headers if the data doesn't have headers.", GH_ParamAccess.list);

            Params.Input.Last().Optional = true;

            pManager.AddIntegerParameter("Empty Entry Handling", "EEH?", "Determine what should Pancake do when it runs into a fully-empty entry.\r\n" +
                "Right-click for more information.\r\n" +
                "0: Keep empty entries; 1: Omit empty entries; 2: Stop at the first encounter. By default 1.", GH_ParamAccess.item, 1);

            var paramEEH = Params.Input.Last() as Param_Integer;
            paramEEH.AddNamedValue("Keep", 0);
            paramEEH.AddNamedValue("Omit", 1);
            paramEEH.AddNamedValue("Stop at first", 2);

            pManager.AddIntegerParameter("Empty Header Handling", "EHH?", "Determine what should Pancake do when it runs into an empty header.\r\n" +
                "Right-click for more information.\r\n" +
                "0: Keep empty entries; 1: Omit empty entries. By default 1.", GH_ParamAccess.item, 1);

            var paramEHH = Params.Input.Last() as Param_Integer;
            paramEHH.AddNamedValue("Keep", 0);
            paramEHH.AddNamedValue("Omit", 1);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Association", "A", "Cell content", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            if (!ShouldRunSolveInstance())
                return;

            GooSheet gooSheet = default;
            GooCellRangeReference gooReferences = default;
            int option = 0;
            bool rowFirst = true;
            bool hasHeader = true;
            bool useNamedHeader = false;
            List<string> headers = new();

            const int EMPTY_KEEP = 0;
            const int EMPTY_OMIT = 1;
            const int EMPTY_STOP = 2;

            var emptyOption = EMPTY_OMIT;
            var emptyHeaderOption = EMPTY_OMIT;

            DA.GetData(0, ref gooSheet);
            DA.GetData(1, ref gooReferences);
            DA.GetData(2, ref option);
            DA.GetData(3, ref rowFirst);
            DA.GetData(4, ref hasHeader);
            DA.GetData(6, ref emptyOption);
            DA.GetData(7, ref emptyHeaderOption);

            if (!hasHeader)
            {
                if (useNamedHeader = DA.GetDataList(5, headers))
                {
                    NotifyForDuplicatedHeaders(headers);
                }
            }
            else
            {
                useNamedHeader = true;
            }

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

            var assocs = new List<IGH_Goo>();
            var id = 0;
            foreach (var branch in cellPositions)
            {
                if (hasHeader && id == 0)
                {
                    foreach (var cref in branch)
                    {
                        object content = null;

                        try
                        {
                            var cell = sheet.GetRow(cref.RowId)?.GetCell(cref.ColumnId);
                            if (cell is null)
                            {
                                content = null;
                            }
                            else
                            {
                                if (emptyHeaderOption != EMPTY_OMIT || !cell.IsEmpty())
                                {
                                    if (!CellAccessUtility.TryGetCellContent(cell, CellTypeHint.Automatic, out content))
                                        content = null;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                        }

                        headers.Add(content?.ToString());
                    }

                    NotifyForDuplicatedHeaders(headers);

                    ++id;
                    continue;
                }

                var assoc = new AssocWrapper();
                var id2 = 0;

                var emptyRow = true;

                foreach (var cref in branch)
                {
                    object content = null;

                    try
                    {
                        var cell = sheet.GetRow(cref.RowId)?.GetCell(cref.ColumnId);
                        if (cell is null)
                        {
                            content = null;
                        }
                        else
                        {
                            if (!cell.IsEmpty())
                                emptyRow = false;

                            if (!CellAccessUtility.TryGetCellContent(cell, hint, out content))
                                content = null;
                        }
                    }
                    catch (Exception ex)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Failed to access {cref}. Replaced with null.");
                    }

                    if (useNamedHeader)
                    {
                        if (id2 >= headers.Count)
                        {
                            AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Not enough headers.");

                            assoc.Add(content);
                        }
                        else
                        {
                            if (headers[id2] is null && emptyHeaderOption == EMPTY_OMIT)
                            {
                            }
                            else
                            {
                                assoc.Add(headers[id2], content);
                            }
                        }
                    }
                    else
                    {
                        assoc.Add(content);
                    }

                    ++id2;
                }

                if (emptyRow)
                {
                    if (emptyOption == EMPTY_OMIT)
                    {
                        assoc = null;
                    }
                    else if (emptyOption == EMPTY_STOP)
                    {
                        break;
                    }
                }

                if (assoc != null)
                    assocs.Add(assoc.AssocObject);

                ++id;
            }

            DA.SetDataList(0, assocs);
        }

        private void NotifyForDuplicatedHeaders(IEnumerable<string> headers)
        {
            if (headers.HasDuplicates())
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Duplicated header names found. Later you may access duplicated named values with indexes only.");
        }
        public override GH_Exposure Exposure => CalculateExposure(GH_Exposure.secondary);

        protected override Bitmap Icon => ComponentIcons.CellToAssoc;
    }
}
