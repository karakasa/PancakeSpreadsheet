using Grasshopper.Kernel;
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
    public class psSheetList : PancakeComponent
    {
        protected override string ComponentCategory => PancakeComponent.CategorySheet;
        public override Guid ComponentGuid => new Guid("{5D91740B-ED1B-4714-9F07-44924FDA3F84}");

        protected override string ComponentName => "Sheet List";

        protected override string ComponentDescription => "Get the list of sheets inside a spreadsheet.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Spreadsheet", "SP", "Spreadsheet object", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Sheets", "S", "Sheet names", GH_ParamAccess.list);
            pManager.AddIntegerParameter("Sheet visbilities", "V", "Whether sheet is hidden.\r\n0: visible; 1: hidden; 2: very hidden", GH_ParamAccess.list);
            pManager.AddIntegerParameter("Active Sheet Index", "A", "Index of the active sheet", GH_ParamAccess.item);
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

            var count = wb.NumberOfSheets;
            var listNames = new List<string>(count);
            var listVisibilities = new List<int>(count);

            for (var i = 0; i < count; i++)
            {
                var name = wb.GetSheetName(i);
                var visibility = 0;

                if (wb.IsSheetHidden(i))
                    visibility = 1;
                if (wb.IsSheetVeryHidden(i))
                    visibility = 2;

                listNames.Add(name);
                listVisibilities.Add(visibility);
            }

            DA.SetDataList(0, listNames);
            DA.SetDataList(1, listVisibilities);

            if (count == 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "No sheets in the spreadsheet object.");

                DA.SetData(2, -1);

                return;
            }

            try
            {
                DA.SetData(2, wb.ActiveSheetIndex);
            }
            catch
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Active sheet is unknown. Something might be wrong.");

                DA.SetData(2, -1);
            }
        }
        protected override Bitmap Icon => ComponentIcons.SheetList;
    }
}
