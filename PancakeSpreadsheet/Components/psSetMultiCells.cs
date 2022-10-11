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
    public class psSetMultiCells : PancakeComponent
    {
        public override Guid ComponentGuid => new("{AAEAC8ED-AA25-448A-997A-F4494D03AB88}");

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
            pManager.AddBooleanParameter("Auto Extend", "E?", AutoExtendParamDesc, GH_ParamAccess.item, true);
        }

        internal const string AutoExtendParamDesc = "True to extend the cell range to ensure all data is written.\r\nFalse to revert to the old behavior, excess data will be truncated.";

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
            bool autoExtend = true;

            DA.GetData(0, ref gooSheet);
            DA.GetData(1, ref gooReferences);
            DA.GetData(2, ref option);
            DA.GetDataTree<IGH_Goo>(3, out var data);
            DA.GetData(4, ref rowFirst);
            DA.GetData(5, ref autoExtend);

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

            var hint = CellAccessUtility.GetHint(option);
            var crange = gooReferences.Value;

            Features.ActualWriteData(sheet, crange, rowFirst, data, autoExtend, hint);

            DA.SetData(0, gooSheet);
        }
        protected override string ComponentCategory => PancakeComponent.CategoryCellContent;
        protected override Bitmap Icon => ComponentIcons.SetCellRange;
    }
}
