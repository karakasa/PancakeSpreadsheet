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
    public class psUnmergeCellRange : PancakeComponent
    {
        public override Guid ComponentGuid => new("{0DC7F94E-BBBB-4D03-938E-6359D3C7A1AD}");

        protected override string ComponentName => "Unmerge Cell Range";

        protected override string ComponentDescription => "Unmerge cell ranges in a sheet.";

        protected override string ComponentCategory => PancakeComponent.CategoryCellStyle;

        public override GH_Exposure Exposure => GH_Exposure.secondary;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Indexes", "I", "Indexes of merged range to unmerge. The index is from 'Get Merged Cell Range' component.", GH_ParamAccess.list);
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

            var list = new List<int>();
            DA.GetDataList(0, list);

            if(list.Count == 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Empty list.");
                return;
            }

            var num = sheet.NumMergedRegions;
            foreach (var it in list)
            {
                if (it < 0 || it >= num)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Invalid index {it}");
                    return;
                }
            }

            sheet.RemoveMergedRegions(list);

            DA.SetData(0, gooSheet);
        }
        protected override Bitmap Icon => ComponentIcons.UnmergeCell;
    }
}
