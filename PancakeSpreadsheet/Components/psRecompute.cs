using Grasshopper.Kernel;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PancakeSpreadsheet.Params;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public class psRecompute : PancakeComponent
    {
        public override Guid ComponentGuid => new("{5BECEC59-0199-4432-9FE6-F16B0C267B86}");

        protected override string ComponentName => "Recompute";

        protected override string ComponentDescription => "Force a spreadsheet to recompute and update all formula cells.\r\n" +
            "Necessary if you want to retrieve the result of formulas after modifying cell contents.\r\n" +
            "Spreadsheet will always be recomputed during saving.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Spreadsheet", "SP", "Spreadsheet object", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Spreadsheet", "SP", "Spreadsheet object", GH_ParamAccess.item);
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

            BaseFormulaEvaluator.EvaluateAllFormulaCells(wb);

            DA.SetData(0, goo);
        }
        protected override string ComponentCategory => PancakeComponent.CategorySpreadsheet;

        public override GH_Exposure Exposure => GH_Exposure.secondary;
        protected override Bitmap Icon => ComponentIcons.Recompute;
    }
}
