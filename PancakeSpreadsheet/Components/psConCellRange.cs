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
    public class psConCellRange : PancakeComponent
    {
        protected override string ComponentCategory => PancakeComponent.CategoryCellContent;
        public override Guid ComponentGuid => new("{CD51C199-E837-1479-815A-2B6BF31843ED}");

        protected override string ComponentName => "Construct Cell Range";

        protected override string ComponentDescription => "Construct a cell reference by row and column index.\r\n" +
            "You may also supply text notations for cell range reference inputs, such as A1:B2.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddParameter(new ParamCellReference(), "Starting Cell", "C(S)", "Cell at left-top corner", GH_ParamAccess.item);
            pManager.AddParameter(new ParamCellReference(), "Ending Cell", "C(E)", "Cell at right-bottom corner", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddParameter(new ParamCellRangeReference(), "Cell Range Reference", "CR", "Cell range reference", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooCellReference gooCref1 = null, gooCref2 = null;

            DA.GetData(0, ref gooCref1);
            DA.GetData(1, ref gooCref2);

            var crange = new SimpleCellRange(gooCref1.Value, gooCref2.Value);

            if (!crange.IsValid())
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid input.");
                return;
            }

            DA.SetData(0, crange.AsGoo());
        }

        public override GH_Exposure Exposure => GH_Exposure.tertiary;
        protected override Bitmap Icon => ComponentIcons.ConCRange;
    }
}
