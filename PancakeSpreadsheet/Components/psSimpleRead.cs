using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using NPOI.SS.UserModel;
using Org.BouncyCastle.Asn1.X509;
using PancakeSpreadsheet.NpoiInterop;
using PancakeSpreadsheet.Params;
using PancakeSpreadsheet.Utility;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public class psSimpleRead : PancakeComponent
    {
        public override Guid ComponentGuid => new("{286082A8-96E7-45FC-ACCF-66DCF7DDA1DD}");

        protected override string ComponentName => "Simple Read";

        protected override string ComponentDescription => "One unified component to read a simple xls(x) file.";

        protected override string ComponentCategory => PancakeComponent.CategorySpreadsheet;
        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddParameter(new Param_FilePath(), "File", "F", "File to read", GH_ParamAccess.item);
            pManager.AddTextParameter("Password", "P", "Password to open the spreadsheet. Leave empty for no password", GH_ParamAccess.item);
            pManager.AddGenericParameter("Sheet Identifier", "S", "Sheet name or index (0-based) to read. By default the first is read.", GH_ParamAccess.item);
            pManager.AddGenericParameter("Read Location", "RL", "Where to read. It can be a starting position, or a range, or nothing. By default the entire sheet is read.", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Row first?", "R?", "Whether to format the data as row-first or column-first. By default row-first, that is, one branch per row.", GH_ParamAccess.item, true);
            pManager.AddBooleanParameter("OK", "OK", "OK to read", GH_ParamAccess.item, false);

            Params.Input[1].Optional = true;
            Params.Input[2].Optional = true;
            Params.Input[3].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Data", "D", "Imported data.", GH_ParamAccess.tree);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string filepath = default;
            string password = default;
            IGH_Goo sheetId = default;
            IGH_Goo readLoc = default;
            bool rowFirst = true;
            bool ok = default;

            DA.GetData(0, ref filepath);
            DA.GetData(1, ref password);
            DA.GetData(2, ref sheetId);
            DA.GetData(3, ref readLoc);
            DA.GetData(4, ref rowFirst);
            DA.GetData(5, ref ok);

            if (!ok)
                return;

            var data = SimpleReadData(filepath, password, sheetId, readLoc, rowFirst);
            if (data is not null)
                DA.SetDataTree(0, data);
        }

        private GH_Structure<IGH_Goo> SimpleReadData(string filepath, string password, IGH_Goo sheetId, IGH_Goo readLoc, bool rowFirst)
        {
            if (!Features.ValidateFile(filepath, out var fileLength))
                return null;

            using var stream = Features.PrepareFileStream(filepath, fileLength);
            using var holder = Features.OpenWorkbook(stream, password);

            if (holder is null)
                return null;

            var sheet = Features.GetSheetByIdentifier(holder, sheetId);
            if (sheet is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Cannot find the specific sheet.");
                return null;
            }

            if (!Features.TryDecideReadRange(sheet, readLoc, out var crange))
                return null;

            return Features.ActualReadData(sheet, crange, rowFirst, CellTypeHint.Automatic);
        }

        protected override void AfterSolveInstance()
        {
            WorkbookFactory.SetImportOption(ImportOption.All);

            base.AfterSolveInstance();
        }
    }
}
