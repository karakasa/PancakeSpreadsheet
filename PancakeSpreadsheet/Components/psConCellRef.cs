using Grasshopper.Kernel;
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
    public class psConCellRef : PancakeComponent
    {
        protected override string ComponentCategory => PancakeComponent.CategoryCellContent;
        public override Guid ComponentGuid => new Guid("{44663403-411F-4A44-9AB5-D3097D61F45A}");

        protected override string ComponentName => "Construct Cell Reference";

        protected override string ComponentDescription => "Construct a cell reference by row and column index.\r\n" +
            "You may also supply A1 or R1C1 notations for cell reference inputs.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddIntegerParameter("Row", "R", "Row index (0-based)", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Column", "C", "Column index (0-based)", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddParameter(new ParamCellReference(), "Cell", "C", "Cell reference", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            int rowId = 0, colId = 0;
            DA.GetData(0, ref rowId);
            DA.GetData(1, ref colId);

            if (rowId < 0 || rowId >= 16384)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid row index.");
                return;
            }

            if (colId < 0 || colId >= 1048576)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid column index.");
                return;
            }

            DA.SetData(0, new SimpleCellReference(rowId, colId).AsGoo());
        }

        public override GH_Exposure Exposure => GH_Exposure.tertiary;
        protected override Bitmap Icon => ComponentIcons.ConCref;
    }
}

