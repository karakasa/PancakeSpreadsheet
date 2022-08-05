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
    public class psAssocToCells : PancakeComponentRequireCore
    {
        const int HEADER_WRITE = 0;
        const int HEADER_DONT_WRITE = 1;
        const int HEADER_WRITE_MATCHED = 2;
        public override Guid ComponentGuid => new("{CDEAC8ED-AA25-448A-997A-F4494D03AB95}");

        protected override string ComponentName => "Assoc to Cells";

        protected override string ComponentDescription => "Set the content of a range of cells, based on data from associations.\r\n" +
            "This component requires the core Pancake experience.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
            pManager.AddParameter(new ParamCellRangeReference(), "Cell Range", "CR", "Cell range reference", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Type", "T", "Cell type for content retrieval\r\nRight-click for more options.", GH_ParamAccess.item, 0);
            this.AddCellTypeHintValues();
            pManager.AddGenericParameter("Associations", "A", "Content to set", GH_ParamAccess.list);
            pManager.AddBooleanParameter("Row First", "R?", "True for row-first data organization; false for column-first.", GH_ParamAccess.item, true);

            pManager.AddTextParameter("Name Filters", "NF", "If set, only values with names in this list are written out.", GH_ParamAccess.list);
            Params.Input.Last().Optional = true;

            pManager.AddBooleanParameter("Clear First", "CF?", "Whether to clear the content of the cell range before writing out. False by default.\r\n" +
                "If Header Handling is set to 'Write matched headers', headers are preserved.", GH_ParamAccess.item, false);
            pManager.AddIntegerParameter("Header Handling", "H?", "Control how names in associations are written into the Excel.\r\n" +
                "Right-click for more information.", GH_ParamAccess.item, 0);
            var paramHeaderHandling = Params.Input.Last() as Param_Integer;

            paramHeaderHandling.AddNamedValue("0: Write headers", HEADER_WRITE);
            paramHeaderHandling.AddNamedValue("1: Don't write headers", HEADER_DONT_WRITE);
            paramHeaderHandling.AddNamedValue("2: Write matched headers only", HEADER_WRITE_MATCHED);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            if (!ShouldRunSolveInstance())
                return;

            GooSheet gooSheet = default;
            GooCellRangeReference gooReferences = default;
            int option = 0;
            bool rowFirst = true;
            bool clearFirst = false;
            int headerHandling = 0;

            var rawData = new List<IGH_Goo>();
            var nameFilters = new List<string>();

            var nameFilterHashset = default(HashSet<string>);

            DA.GetData(0, ref gooSheet);
            DA.GetData(1, ref gooReferences);
            DA.GetData(2, ref option);
            DA.GetDataList(3, rawData);
            DA.GetData(4, ref rowFirst);

            var useNameFilters = DA.GetDataList(5, nameFilters);
            if (nameFilters.Count == 0)
            {
                useNameFilters = false;
            }
            else
            {
                nameFilterHashset = new HashSet<string>(nameFilters);
            }

            DA.GetData(6, ref clearFirst);
            DA.GetData(7, ref headerHandling);

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

            var assocs = rawData.Select(static r => new AssocWrapper(r)).ToArray();
            var effectiveNames = new List<string>();
            var effectiveNameHashset = new HashSet<string>();

            foreach (var assoc in assocs)
            {
                var names = assoc.Names;
                foreach (var name in names)
                {
                    if (useNameFilters && !nameFilterHashset.Contains(name))
                        continue;

                    if (effectiveNameHashset.Add(name))
                        effectiveNames.Add(name);
                }
            }

            var crange = gooReferences.Value;

            var cellPositions = rowFirst ? crange.EnumerateRowFirst() : crange.EnumerateColumnFirst();

            var hint = CellAccessUtility.GetHint(option);
            var id = -1;

            var assocId = 0;

            foreach (var branch in cellPositions)
            {
                ++id;

                if (id == 0)
                {
                    if (headerHandling == HEADER_WRITE)
                    {
                        var index = 0;
                        foreach (var cref in branch)
                        {
                            var cell = sheet.EnsureCell(cref);
                            if (index < effectiveNames.Count)
                            {
                                cell.SetCellValue(effectiveNames[index]);
                            }
                            else
                            {
                                if (clearFirst)
                                    cell.SetBlank();
                            }

                            ++index;
                        }

                        continue;
                    }
                    else if (headerHandling == HEADER_WRITE_MATCHED)
                    {
                        effectiveNames.Clear();

                        foreach (var cref in branch)
                        {
                            var curName = default(string);

                            try
                            {
                                var cell = sheet.CellAt(cref);
                                if (cell is not null && cell.CellType != CellType.Blank)
                                {
                                    curName = cell.ToString();
                                    if (!effectiveNameHashset.Contains(curName))
                                        curName = default;
                                }
                            }
                            catch
                            {

                            }

                            effectiveNames.Add(curName);
                        }

                        continue;
                    }
                }

                var effectiveNameId = 0;

                foreach (var cref in branch)
                {
                    ICell cell = default;

                    if (clearFirst)
                    {
                        cell = sheet.CellAt(cref);
                        if (cell is not null)
                            cell.SetBlank();
                    }

                    if (assocId >= assocs.Length)
                        continue;

                    if (effectiveNameId >= effectiveNames.Count)
                        continue;

                    var effectiveName = effectiveNames[effectiveNameId];

                    if (effectiveName is null)
                    {
                        ++effectiveNameId;
                        continue;
                    }

                    cell ??= sheet.EnsureCell(cref);
                    if (!CellAccessUtility.TrySetCellContentPrimitive(cell, hint, assocs[assocId].Get(effectiveName)))
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Cannot set the content of cell {cref}. Skipped.");
                    }

                    ++effectiveNameId;
                }

                ++assocId;
            }

            DA.SetData(0, gooSheet);
        }
        protected override string ComponentCategory => PancakeComponent.CategoryCellContent;

        public override GH_Exposure Exposure => CalculateExposure(GH_Exposure.secondary);
        protected override Bitmap Icon => ComponentIcons.AssocToCell;
    }
}
