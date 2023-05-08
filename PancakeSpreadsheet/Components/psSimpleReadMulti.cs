using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using NPOI.SS.UserModel;
using PancakeSpreadsheet.NpoiInterop;
using PancakeSpreadsheet.Params;
using PancakeSpreadsheet.Utility;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public class psSimpleReadMulti : PancakeComponent
    {
        public override Guid ComponentGuid => new("{286082A8-96E7-45FC-ACCF-66DCF7DDA1EE}");

        protected override string ComponentName => "Simple Read XLS(X) Multiple Sheets";

        protected override string ComponentDescription => "One unified component to read multiple sheets from a xls(x) file.";

        protected override string ComponentCategory => PancakeComponent.CategorySpreadsheet;
        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddParameter(new Param_FilePath(), "File", "F", "File to read", GH_ParamAccess.item);
            pManager.AddTextParameter("Password", "P", "Password to open the spreadsheet. Leave empty for no password", GH_ParamAccess.item);
            pManager.AddGenericParameter("Sheet Identifier", "S", "Sheet name or index (0-based) to read.", GH_ParamAccess.list);
            pManager.AddGenericParameter("Read Location", "RL", "Where to read. It can be a starting position, or a range, or nothing. By default the entire sheet is read.", GH_ParamAccess.list);
            pManager.AddBooleanParameter("Row first?", "R?", "Whether to format the data as row-first or column-first. By default row-first, that is, one branch per row.", GH_ParamAccess.item, true);
            pManager.AddBooleanParameter("OK", "OK", "OK to read", GH_ParamAccess.item, false);

            Params.Input[1].Optional = true;
            Params.Input[2].Optional = true;
            Params.Input[3].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Data", "D", "Imported data.", GH_ParamAccess.tree);
            pManager.AddTextParameter("Sheet names", "N", "Effective sheet names.", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string filepath = default;
            string password = default;
            var sheetId = new List<IGH_Goo>();
            var readLoc = new List<IGH_Goo>();
            bool rowFirst = true;
            bool ok = default;

            DA.GetData(0, ref filepath);
            DA.GetData(1, ref password);
            DA.GetDataList(2, sheetId);
            DA.GetDataList(3, readLoc);
            DA.GetData(4, ref rowFirst);
            DA.GetData(5, ref ok);

            if (!ok)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Wire a True value to the OK input, to read the file.");
                return;
            }

            var data = SimpleReadDataMultipleSheets(filepath, password, sheetId, readLoc, rowFirst, out var names);
            if (data is not null)
            {
                DA.SetDataTree(0, data);
                DA.SetDataList(1, names);
            }
        }

        private GH_Structure<IGH_Goo> SimpleReadDataMultipleSheets(
            string filepath, 
            string password, 
            List<IGH_Goo> sheetIds, 
            List<IGH_Goo> readLocs,
            bool rowFirst,
            out List<string> names
            )
        {
            names = null;

            if (!Features.ValidateFile(filepath, out var fileLength))
                return null;

            using var stream = Features.PrepareFileStream(filepath, fileLength);
            using var holder = Features.OpenWorkbook(stream, password);

            if (holder is null)
                return null;

            var tree = new GH_Structure<IGH_Goo>();
            var index = 0;

            if (sheetIds.Count == 0)
            {
                var count = holder.Workbook.NumberOfSheets;
                for (var i = 0; i < count; i++)
                {
                    if (holder.Workbook.IsSheetHidden(i)
                        || holder.Workbook.IsSheetVeryHidden(i))
                        continue;

                    var name = holder.Workbook.GetSheetName(i);
                    if (name is not null)
                        sheetIds.Add(new GH_String(name));
                }
            }

            names = new List<string>();

            foreach (var sheetId in sheetIds)
            {
                var path = new GH_Path(index);
                var sheet = Features.GetSheetByIdentifier(holder, sheetId);
                if (sheet is null)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Cannot find the specific sheet {sheetId?.ToString() ?? null}.");
                    tree.EnsurePath(path);
                    names.Add(null);
                    continue;
                }

                var readLoc = (readLocs.Count > index) ? readLocs[index] : null;

                if (!Features.TryDecideReadRange(sheet, readLoc, out var crange))
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Cannot decide the read range of the specific sheet {sheetId?.ToString() ?? null}.");
                    tree.EnsurePath(path);
                    names.Add(null);
                    continue;
                }

                names.Add(sheet.SheetName);
                Features.ActualReadData(sheet, crange, rowFirst, CellTypeHint.Automatic, tree, path);

                ++index;
            }

            return tree;
        }

        protected override void AfterSolveInstance()
        {
            WorkbookFactory.SetImportOption(ImportOption.All);

            base.AfterSolveInstance();
        }
        protected override Bitmap Icon => ComponentIcons.SimpleRead;
    }
}

