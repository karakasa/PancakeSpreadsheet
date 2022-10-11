using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PancakeSpreadsheet.NpoiInterop;
using PancakeSpreadsheet.Params;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public class psSaveSpreadsheet : PancakeComponent
    {
        public override Guid ComponentGuid => new("{B66C867B-05C1-4519-9B6B-FDA99127D77E}");

        protected override string ComponentName => "Save Spreadsheet";

        protected override string ComponentDescription => "Save a spreadsheet.\r\n" +
            "Please use the 'Wait Until' component to postpone this component and ensure all modifications are done BEFORE saving.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Spreadsheet", "SP", "Spreadsheet object", GH_ParamAccess.item);
            pManager.AddParameter(new Param_FilePath(), "File", "F", "Location to save", GH_ParamAccess.item);
            pManager.AddTextParameter("Password", "P", "Not supported yet.", GH_ParamAccess.item, string.Empty);
            pManager.AddBooleanParameter("Overwrite", "OW", "Whether to overwrite existing files", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("OK", "OK", "Whether to proceed", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddBooleanParameter("OK", "OK", "Whether the operation is successful", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string filepath = default;
            string password = default;
            bool overwrite = false;
            bool ok = false;

            GooSpreadsheet goo = default;

            DA.GetData(0, ref goo);

            var holder = goo?.Value;

            var wb = holder.Workbook;

            if (wb is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid spreadsheet.");
                return;
            }

            DA.GetData(1, ref filepath);
            DA.GetData(2, ref password);
            DA.GetData(3, ref overwrite);
            DA.GetData(4, ref ok);

            if (!ok)
                return;

            if (!string.IsNullOrEmpty(password))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Saving with password is not supported yet.");
                return;
            }

            Features.WriteToFile(filepath, holder, overwrite);
        }
        protected override string ComponentCategory => PancakeComponent.CategorySpreadsheet;
        protected override Bitmap Icon => ComponentIcons.SaveFile;
    }
}
