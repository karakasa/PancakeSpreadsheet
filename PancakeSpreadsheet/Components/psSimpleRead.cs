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
            if (!ValidateFile(filepath, out var fileLength))
                return null;

            using var stream = PrepareFileStream(filepath, fileLength);
            using var holder = OpenWorkbook(stream, password);

            if (holder is null)
                return null;

            var sheet = sheetId is null
                ? holder.Workbook.GetSheetAt(0)
                : ConversionUtility.TryGetIndexOrName(sheetId, out var index, out var name)
                    switch
                {
                    IndexNameState.Index => holder.Workbook.GetSheetAt(index),
                    IndexNameState.Name => holder.Workbook.GetSheet(name),
                    _ => null
                };

            if (sheet is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Cannot find the specific sheet.");
                return null;
            }

            if (!TryDecideReadRange(sheet, readLoc, out var crange))
                return null;

            return ActualReadData(sheet, crange, rowFirst, CellTypeHint.Automatic);
        }

        private GH_Structure<IGH_Goo> ActualReadData(ISheet sheet, SimpleCellRange crange, bool rowFirst, CellTypeHint hint)
        {
            var cellPositions = rowFirst ? crange.EnumerateRowFirst() : crange.EnumerateColumnFirst();

            var tree = new GH_Structure<IGH_Goo>();
            var id = 0;
            foreach (var branch in cellPositions)
            {
                var list = new List<IGH_Goo>();

                foreach (var cref in branch)
                {
                    try
                    {
                        object content;

                        var cell = sheet.GetRow(cref.RowId)?.GetCell(cref.ColumnId);
                        if (cell is null)
                        {
                            content = null;
                        }
                        else
                        {
                            if (!CellAccessUtility.TryGetCellContent(cell, hint, out content))
                                content = null;
                        }

                        list.Add(GH_Convert.ToGoo(content));
                    }
                    catch (Exception ex)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Failed to access {cref}. Replaced with null.");
                        list.Add(null);
                    }
                }

                tree.EnsurePath(id).AddRange(list);
                ++id;
            }

            return tree;
        }

        private bool TryDecideReadRange(ISheet sheet, IGH_Goo readLoc, out SimpleCellRange crange)
        {
            crange = GetSheetRange(sheet);
            if (!crange.IsValid())
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Empty sheet.");
                return false;
            }

            if (readLoc is not null)
            {
                var state = CellAccessUtility.TryGetCellData(readLoc, out var expectedCref, out var expectedCrange);
                if (state is CellAccessUtility.CellDataType.Invalid)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid read location.");
                    return false;
                }

                if (state is CellAccessUtility.CellDataType.CellRef)
                {
                    if (crange.Contains(expectedCref))
                    {
                        crange.StartCell = expectedCref;
                    }
                    else
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Read location is not within the range of sheet. Ignored.");
                    }
                }
                else if (state is CellAccessUtility.CellDataType.CellRange)
                {
                    crange = expectedCrange;
                }
            }

            return true;
        }
        private static SimpleCellRange GetSheetRange(ISheet sheet)
        {
            var firstRowIndex = sheet.FirstRowNum;
            var lastRowIndex = sheet.LastRowNum;

            var firstColIndex = int.MaxValue;
            var lastColIndex = int.MinValue;

            for (var rowIndex = firstRowIndex; rowIndex <= lastRowIndex; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row is null)
                    continue;

                if (row.FirstCellNum < firstColIndex)
                    firstColIndex = row.FirstCellNum;

                if (row.LastCellNum - 1 > lastColIndex)
                    lastColIndex = row.LastCellNum - 1;
            }

            return new SimpleCellRange(new(firstRowIndex, firstColIndex), new(lastRowIndex, lastColIndex));
        }
        private WorkbookHolder OpenWorkbook(Stream stream, string password)
        {
            WorkbookHolder holder = default;
            WorkbookFactory.SetImportOption(ImportOption.SheetContentOnly);

            try
            {
                holder = WorkbookHolder.Create(stream, password);
            }
            catch (NotSupportedException nsEx)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, nsEx.Message);
                return null;
            }
            catch (IOException ioEx)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, ioEx.InnerException?.Message ?? ioEx.Message);
                return null;
            }

            return holder;
        }

        private Stream PrepareFileStream(string filepath, long fileLength)
        {
            const long InPlaceFileSizeThreshold = 1L * 1024 * 1024 * 1024; // 1GB

            Stream stream;

            if (fileLength > InPlaceFileSizeThreshold)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "File is too large. Pancake will open the file in-place. This may cause issues.");
                stream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }
            else
            {
                stream = new MemoryStream((int)fileLength);
                using var fileStream = StaticExtensions.OpenFile(filepath);

                fileStream.CopyTo(stream);
            }

            return stream;
        }
        private bool ValidateFile(string filepath, out long fileLength)
        {
            fileLength = 0;

            if (!File.Exists(filepath))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "File not found.");
                return false;
            }

            var extension = Path.GetExtension(filepath).ToLowerInvariant();
            switch (extension)
            {
                case ".csv":
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "CSV file is unsupported. Use Pancake for CSV import & export.");
                    return false;
                case ".xlsb":
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "XLSB file is unsupported.");
                    return false;
            }

            fileLength = new FileInfo(filepath).Length;
            if (fileLength < 32)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Empty file.");
                return false;
            }

            return true;
        }

        protected override void AfterSolveInstance()
        {
            WorkbookFactory.SetImportOption(ImportOption.All);

            base.AfterSolveInstance();
        }
    }
}
