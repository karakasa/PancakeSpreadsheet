using Grasshopper.Kernel;
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
    public class psCreateSheet : PancakeComponent
    {
        protected override string ComponentCategory => PancakeComponent.CategorySheet;
        public override Guid ComponentGuid => new("{0C5A40AA-C442-4C1A-9932-0C2B18EE02A5}");

        protected override string ComponentName => "Create Sheet";

        protected override string ComponentDescription => "Create a sheet by name.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Spreadsheet", "SP", "Spreadsheet object", GH_ParamAccess.item);
            pManager.AddTextParameter("Name", "N", "Name of the sheet", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooSpreadsheet goo = default;

            DA.GetData(0, ref goo);

            var wb = goo?.Value?.Workbook;

            if (wb is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid spreadsheet.");
                return;
            }

            string name = default;

            DA.GetData(1, ref name);

            if (string.IsNullOrEmpty(name))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Name cannot be empty.");
                return;
            }

            ISheet sheet = wb.CreateSheet(name);
            DA.SetData(0, sheet.AsGoo());
        }
        protected override Bitmap Icon => ComponentIcons.NewSheet;
        public override GH_Exposure Exposure => GH_Exposure.secondary;
    }
}
