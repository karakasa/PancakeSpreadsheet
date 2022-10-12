using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PancakeSpreadsheet.NpoiInterop;
using PancakeSpreadsheet.Params;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public class psSimpleWrite : PancakeComponent
    {
        public override Guid ComponentGuid => new("{7F7652DF-0286-4F77-ADDF-F4F734D3D297}");

        protected override string ComponentName => "Simple Write XLS";

        protected override string ComponentDescription => "One unified component to write a simple xls(x) file.";

        protected override string ComponentCategory => PancakeComponent.CategorySpreadsheet;
        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddParameter(new Param_FilePath(), "File", "F", "File to read", GH_ParamAccess.item);
            pManager.AddGenericParameter("Sheet Identifier", "S", "Sheet name or index (0-based) to write.", GH_ParamAccess.item);
            pManager.AddParameter(new ParamCellReference(), "Write Position", "WP", "Where to start writing. By default the top-left corner is used.", GH_ParamAccess.item);
            pManager.AddGenericParameter("Data", "D", "Data to write", GH_ParamAccess.tree);
            pManager.AddBooleanParameter("Row first?", "R?", "Whether to format the data as row-first or column-first. By default row-first, that is, one branch per row.", GH_ParamAccess.item, true);
            pManager.AddBooleanParameter("Resize Column?", "RC?", "Resize all columns to fit content. By default true.", GH_ParamAccess.item, true);
            pManager.AddBooleanParameter("Create New?", "CN?", "If the output file exists, true to delete the file first, false to write onto the original file.", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("OK", "OK", "OK to write", GH_ParamAccess.item, false);

            Params.Input[1].Optional = true;
            Params.Input[2].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddBooleanParameter("OK", "OK", "Whether the operation is successful.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string filepath = default;
            IGH_Goo sheetId = default;
            GooCellReference writePosition = default;
            bool rowFirst = true;
            bool resizeCol = true;
            bool createNew = false;
            bool ok = default;

            DA.GetData(0, ref filepath);
            DA.GetData(1, ref sheetId);
            DA.GetData(2, ref writePosition);
            DA.GetDataTree<IGH_Goo>(3, out var dataTree);
            DA.GetData(4, ref rowFirst);
            DA.GetData(5, ref resizeCol);
            DA.GetData(6, ref createNew);
            DA.GetData(7, ref ok);

            if (!ok)
                return;

            var position = writePosition == null ? new SimpleCellReference(0, 0) : writePosition.Value;

            if (!Features.ValidateFile(filepath, out _, true))
                return;

            var isOOXMLFormat = !Path.GetExtension(filepath).Equals(".xls", StringComparison.OrdinalIgnoreCase);

            if (createNew && File.Exists(filepath))
            {
                try
                {
                    File.Delete(filepath);
                    if (File.Exists(filepath))
                        throw new IOException();
                }
                catch
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Fail to erase the existing file.");
                    return;
                }
            }

            WorkbookHolder holder;

            var useExistingFile = false;

            if (File.Exists(filepath))
            {
                var stream = Features.PrepareMemoryStream(filepath);
                holder = Features.OpenWorkbook(stream, string.Empty);
                useExistingFile = true;
            }
            else
            {
                IWorkbook wb = isOOXMLFormat ? new XSSFWorkbook() : new HSSFWorkbook();
                holder = WorkbookHolder.Create(null, wb);
            }

            if (holder is null)
                return;

            const string DEFAULT_SHEET_NAME = "Sheet1";

            try
            {
                ISheet sheet = null;

                var state = ConversionUtility.TryGetIndexOrName(sheetId, out var index, out var name);

                if (useExistingFile)
                    sheet = Features.GetSheetByIdentifier(holder, sheetId);

                if (sheet is null)
                {
                    switch (state)
                    {
                        case IndexNameState.Name when !string.IsNullOrEmpty(name):
                            sheet = holder.Workbook.CreateSheet(name);
                            break;
                        case IndexNameState.Index:
                            AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Unable to create a sheet with specific index. Revert to the default sheet name.");
                            sheet = holder.Workbook.CreateSheet(DEFAULT_SHEET_NAME);
                            break;
                        default:
                            if (sheetId is null)
                            {
                                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "A default sheet name is used.");
                                sheet = holder.Workbook.CreateSheet(DEFAULT_SHEET_NAME);
                            }
                            else
                            {
                                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid sheet identifier");
                                return;
                            }
                            break;
                    }
                }

                Features.ActualWriteData(sheet, new SimpleCellRange(position), rowFirst, dataTree, true, CellTypeHint.Automatic);

                if (resizeCol)
                    Features.ResizeAll(sheet);

                Features.WriteToFile(filepath, holder, true);
            }
            finally
            {
                holder?.Dispose();
            }
        }
        protected override Bitmap Icon => ComponentIcons.SimpleWrite;
    }
}
