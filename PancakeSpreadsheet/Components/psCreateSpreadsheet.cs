using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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
    public class psCreateSpreadsheet : PancakeComponentWithLongstandingRsrc
    {
        protected override string ComponentCategory => PancakeComponent.CategorySpreadsheet;
        public override Guid ComponentGuid => new("{A66C867B-05C1-4519-9B6B-FDA99127D77E}");

        protected override string ComponentName => "Create Spreadsheet";

        protected override string ComponentDescription => "Create a spreadsheet with specific settings.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddIntegerParameter("Type", "T", "Type of the spreadsheet. Right-click for more options.\r\nPassword-protected ", GH_ParamAccess.item, 0);

            var paramType = Params.Input.Last() as Param_Integer;
            paramType.AddNamedValue("OOXML (xlsx)", 0);
            paramType.AddNamedValue("Compound (xls)", 1);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Spreadsheet", "SP", "Spreadsheet object", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            var type = 0;

            DA.GetData(0, ref type);

            IWorkbook wb;

            switch (type)
            {
                case 0:
                    wb = new XSSFWorkbook();
                    break;
                case 1:
                    wb = new HSSFWorkbook();
                    break;
                default:
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Unknown type.");
                    return;
            }

            var holder = WorkbookHolder.Create(null, wb);

            MonitorResource(holder);

            DA.SetData(0, holder.AsGoo());
        }
        protected override Bitmap Icon => ComponentIcons.NewFile;
    }
}
