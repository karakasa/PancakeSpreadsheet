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
    public class psDeconCellRange : PancakeComponent
    {
        public override GH_Exposure Exposure => GH_Exposure.tertiary;
        public override Guid ComponentGuid => new Guid("{95C1285A-4512-4D19-AAE9-C6020D97CF68}");

        protected override string ComponentName => "Deconstruct Cell Range";

        protected override string ComponentDescription => "Decompose a cell reference into row and column index.";

        protected override string ComponentCategory => CategoryCellContent;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddParameter(new ParamCellRangeReference(), "Cell Range", "CR", "Cell range reference", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddParameter(new ParamCellReference(), "Starting Cell", "C(S)", "Cell at left-top corner", GH_ParamAccess.item);
            pManager.AddParameter(new ParamCellReference(), "Ending Cell", "C(E)", "Cell at right-bottom corner", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooCellRangeReference cref = default;
            if (!DA.GetData(0, ref cref) || !cref.Value.IsValid())
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid input.");
                return;
            }

            DA.SetData(0, cref.Value.StartCell.AsGoo());
            DA.SetData(1, cref.Value.EndCell.AsGoo());
        }
        protected override Bitmap Icon => ComponentIcons.DeconCRange;
    }
}
