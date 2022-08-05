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
    public class psSetCellStyle : PancakeComponent
    {
        protected override string ComponentCategory => PancakeComponent.CategoryCellStyle;
        public override Guid ComponentGuid => new("{C8922373-306F-40C1-A54D-11A0FD3D49A7}");

        protected override string ComponentName => "Set Cell Style";

        protected override string ComponentDescription => "Set the style of one or more cells.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
            pManager.AddGenericParameter("Cell(s)", "C/CR", "Cell reference or cell range reference", GH_ParamAccess.item);
            pManager.AddTextParameter("Style Name", "N", "Names of style to be set", GH_ParamAccess.list);
            pManager.AddGenericParameter("Style Value", "V", "Values of style to be set", GH_ParamAccess.list);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooSheet gooSheet = default;

            DA.GetData(0, ref gooSheet);

            var sheet = gooSheet?.Value;

            if (sheet is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid sheet.");
                return;
            }

            object gooTarget = null;
            SimpleCellReference[] appliedCells;

            DA.GetData(1, ref gooTarget);

            if (gooTarget is GooCellReference cref)
            {
                appliedCells = new[] { cref.Value };
            }
            else if (gooTarget is GooCellRangeReference crange)
            {
                appliedCells = crange.Value.Enumerate().ToArray();
            }
            else
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Input must be a cell reference or a cell range reference");
                return;
            }

            List<string> names = new();
            List<IGH_Goo> values = new();

            DA.GetDataList(2, names);
            DA.GetDataList(3, values);

            if (names.Count == 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Must set at least 1 style.");
                return;
            }

            if (names.Count != values.Count)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Length of names and values mismatch.");
                return;
            }

            var propertyDictionary = new Dictionary<string, object>();

            for (var i = 0; i < names.Count; i++)
                propertyDictionary[names[i]] = values[i].AsPrimitive();

            foreach (var cellRef in appliedCells)
            {
                var cell = sheet.EnsureCell(cellRef.RowId, cellRef.ColumnId);

                CellUtil.SetCellStyleProperties(cell, propertyDictionary);
            }

            DA.SetData(0, gooSheet);
        }
        protected override Bitmap Icon => ComponentIcons.SetStyle;
    }
}
