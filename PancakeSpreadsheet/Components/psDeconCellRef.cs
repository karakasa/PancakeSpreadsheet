using Grasshopper.Kernel;
using PancakeSpreadsheet.Params;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public class psDeconCellRef : PancakeComponent
    {
        public override GH_Exposure Exposure => GH_Exposure.tertiary;
        public override Guid ComponentGuid => new Guid("{95C1285A-4512-4D19-AAE9-C6020D97CF67}");

        protected override string ComponentName => "Deconstruct Cell Reference";

        protected override string ComponentDescription => "Decompose a cell reference into row and column index.";

        protected override string ComponentCategory => CategoryCellContent;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddParameter(new ParamCellReference(), "Cell", "C", "Cell reference", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddIntegerParameter("Row", "R", "Row index (0-based)", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Column", "C", "Column index (0-based)", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooCellReference cref = default;
            if (!DA.GetData(0, ref cref) || !cref.Value.IsValid())
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid input.");
                return;
            }

            DA.SetData(0, cref.Value.RowId);
            DA.SetData(1, cref.Value.ColumnId);
        }
        protected override Bitmap Icon => ComponentIcons.DeconCref;
    }
}
