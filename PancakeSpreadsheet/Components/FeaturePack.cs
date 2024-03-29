﻿using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Grasshopper.Kernel;
using NPOI.SS.UserModel;
using PancakeSpreadsheet.NpoiInterop;
using PancakeSpreadsheet.Utility;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Bcpg.Sig;
using NPOI.SS.Formula;
using static ICSharpCode.SharpZipLib.Zip.FastZip;
using System.Runtime.InteropServices.WindowsRuntime;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Formula.Functions;

namespace PancakeSpreadsheet.Components
{
    internal class FeaturePack
    {
        private readonly GH_ActiveObject _docObj;
        public FeaturePack(GH_ActiveObject activeObj)
        {
            _docObj = activeObj;
        }
        public void ResizeAll(ISheet sheet)
        {
            GetSheetRange(sheet, out _, out _, out var firstColIndex, out var lastColIndex);

            for (var colIndex = firstColIndex; colIndex <= lastColIndex; colIndex++)
            {
                try
                {
                    sheet.AutoSizeColumn(colIndex);
                }
                catch
                {
                }
            }
        }
        public bool WriteToFile(string filepath, WorkbookHolder holder, bool overwrite)
        {
            if (!overwrite && File.Exists(filepath))
            {
                _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "The destination file already exists.");
                return false;
            }

            if (overwrite && File.Exists(filepath))
            {
                try
                {
                    File.Delete(filepath);
                }
                catch
                {
                    _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Fail to delete the destination file. It may be opened already by another software.");
                    return false;
                }
            }

            if (holder.IsFileBased && filepath.Equals(holder.BackendFile, StringComparison.InvariantCultureIgnoreCase))
            {
                _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "You cannot overwrite the original file, because the file is opened in-place.");
                return false;
            }

            BaseFormulaEvaluator.EvaluateAllFormulaCells(holder.Workbook);

            using var fs = new FileStream(filepath, FileMode.Create, FileAccess.Write);
            holder.Workbook.Write(fs, true);

            return true;
        }
        public GH_Structure<IGH_Goo> ActualReadData(ISheet sheet, SimpleCellRange crange, bool rowFirst, CellTypeHint hint)
        {
            var tree = new GH_Structure<IGH_Goo>();
            ActualReadData(sheet, crange, rowFirst, hint, tree, null);
            return tree;
        }
        public void ActualReadData(ISheet sheet, SimpleCellRange crange, bool rowFirst, CellTypeHint hint, GH_Structure<IGH_Goo> tree, GH_Path basePath)
        {
            var cellPositions = rowFirst ? crange.EnumerateRowFirst() : crange.EnumerateColumnFirst();

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
                        _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Failed to access {cref}. Replaced with null.");
                        list.Add(null);
                    }
                }

                GH_Path curPath;
                if (basePath is null)
                    curPath = new GH_Path(id);
                else
                    curPath = basePath.AppendElement(id);

                tree.EnsurePath(curPath).AddRange(list);
                ++id;
            }
        }

        public bool TryDecideReadRange(ISheet sheet, IGH_Goo readLoc, out SimpleCellRange crange)
        {
            crange = GetSheetRange(sheet);
            if (!crange.IsValid())
            {
                _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Empty sheet.");
                return false;
            }

            if (readLoc is not null)
            {
                var state = CellAccessUtility.TryGetCellData(readLoc, out var expectedCref, out var expectedCrange);
                if (state is CellAccessUtility.CellDataType.Invalid)
                {
                    _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid read location.");
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
                        _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Read location is not within the range of sheet. Ignored.");
                    }
                }
                else if (state is CellAccessUtility.CellDataType.CellRange)
                {
                    crange = expectedCrange;
                }
            }

            return true;
        }
        public SimpleCellRange GetSheetRange(ISheet sheet)
        {
            GetSheetRange(sheet, out var firstRowIndex, out var lastRowIndex, out var firstColIndex, out var lastColIndex);
            return new SimpleCellRange(new(firstRowIndex, firstColIndex), new(lastRowIndex, lastColIndex));
        }
        public void GetSheetRange(ISheet sheet, out int firstRowIndex, out int lastRowIndex, out int firstColIndex, out int lastColIndex)
        {
            firstRowIndex = sheet.FirstRowNum;
            lastRowIndex = sheet.LastRowNum;

            firstColIndex = int.MaxValue;
            lastColIndex = int.MinValue;

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
        }
        public WorkbookHolder OpenWorkbook(Stream stream, string password)
        {
            WorkbookHolder holder = default;
            WorkbookFactory.SetImportOption(ImportOption.All);

            try
            {
                holder = WorkbookHolder.Create(stream, password);
            }
            catch (NotSupportedException nsEx)
            {
                _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, nsEx.Message);
                return null;
            }
            catch (IOException ioEx)
            {
                _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, ioEx.InnerException?.Message ?? ioEx.Message);
                return null;
            }

            return holder;
        }
        public ISheet GetSheetByIdentifier(WorkbookHolder holder, IGH_Goo sheetId)
        {
            return sheetId is null
                ? holder.Workbook.GetSheetAt(0)
                : ConversionUtility.TryGetIndexOrName(sheetId, out var index, out var name)
                    switch
                {
                    IndexNameState.Index => holder.Workbook.GetSheetAt(index),
                    IndexNameState.Name => holder.Workbook.GetSheet(name),
                    _ => null
                };
        }

        public MemoryStream PrepareMemoryStream(string filepath, long fileLength = -1)
        {
            if (fileLength <= 0)
                fileLength = new FileInfo(filepath).Length;

            var stream = new MemoryStream((int)fileLength);
            using var fileStream = StaticExtensions.OpenFile(filepath);

            fileStream.CopyTo(stream);
            return stream;
        }
        public Stream PrepareFileStream(string filepath, long fileLength)
        {
            const long InPlaceFileSizeThreshold = 1L * 1024 * 1024 * 1024; // 1GB

            Stream stream;

            if (fileLength > InPlaceFileSizeThreshold)
            {
                _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "File is too large. Pancake will open the file in-place. This may cause issues.");
                stream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }
            else
            {
                stream = PrepareMemoryStream(filepath, fileLength);
            }

            return stream;
        }
        public WorkbookHolder OpenOrCreateFile(string filepath, bool alwaysCreateNew, out bool useExistingFile)
        {
            var isOOXMLFormat = !Path.GetExtension(filepath).Equals(".xls", StringComparison.OrdinalIgnoreCase);

            if (alwaysCreateNew && File.Exists(filepath))
            {
                try
                {
                    File.Delete(filepath);
                    if (File.Exists(filepath))
                        throw new IOException();
                }
                catch
                {
                    _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Fail to erase the existing file.");
                    useExistingFile = false;
                    return null;
                }
            }

            WorkbookHolder holder;

            useExistingFile = false;

            if (File.Exists(filepath))
            {
                var stream = PrepareMemoryStream(filepath);
                holder = OpenWorkbook(stream, string.Empty);
                useExistingFile = true;
            }
            else
            {
                IWorkbook wb = isOOXMLFormat ? new XSSFWorkbook() : new HSSFWorkbook();
                holder = WorkbookHolder.Create(null, wb);
            }

            return holder;
        }
        public ISheet OpenOrCreateSheet(WorkbookHolder holder, IGH_Goo sheetId, bool useExistingFile)
        {
            const string DEFAULT_SHEET_NAME = "Sheet1";

            ISheet sheet = null;

            var state = ConversionUtility.TryGetIndexOrName(sheetId, out var index, out var name);

            if (useExistingFile)
                sheet = GetSheetByIdentifier(holder, sheetId);

            if (sheet is null)
            {
                switch (state)
                {
                    case IndexNameState.Name when !string.IsNullOrEmpty(name):
                        sheet = holder.Workbook.CreateSheet(name);
                        break;
                    case IndexNameState.Index:
                        _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Unable to create a sheet with specific index. Revert to the default sheet name.");
                        sheet = holder.Workbook.CreateSheet(DEFAULT_SHEET_NAME);
                        break;
                    default:
                        if (sheetId is null)
                        {
                            _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "A default sheet name is used.");
                            sheet = holder.Workbook.CreateSheet(DEFAULT_SHEET_NAME);
                        }
                        else
                        {
                            _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid sheet identifier");
                            return null;
                        }
                        break;
                }
            }

            return sheet;
        }
        public bool ValidateFile(string filepath, out long fileLength, bool forWrite = false)
        {
            fileLength = 0;

            var extension = Path.GetExtension(filepath).ToLowerInvariant();
            switch (extension)
            {
                case ".csv":
                    _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "CSV file is unsupported. Use Pancake for CSV import & export.");
                    return false;
                case ".xlsb":
                    _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "XLSB file is unsupported.");
                    return false;
            }

            if (forWrite)
                return true;

            if (!File.Exists(filepath))
            {
                _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "File not found.");
                return false;
            }

            fileLength = new FileInfo(filepath).Length;
            if (fileLength < 32)
            {
                _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Empty file.");
                return false;
            }

            return true;
        }

        public void ActualWriteData(
            ISheet sheet,
            SimpleCellRange crange,
            bool rowFirst,
            GH_Structure<IGH_Goo> data,
            bool autoExtend,
            CellTypeHint hint,
            bool ignoreNull = false
            )
        {
            var cntFirstLevel = crange.RowCount;
            var cntSecondLevel = crange.ColumnCount;

            if (!rowFirst)
                StaticExtensions.Swap(ref cntFirstLevel, ref cntSecondLevel);

            var dataCntFirstLevel = data.PathCount;
            var dataCntSecondLevel = autoExtend ? data.Branches.Max(branch => branch.Count)
                : data.Branches.Min(branch => branch.Count);

            if (autoExtend)
            {
                if (dataCntFirstLevel > cntFirstLevel)
                {
                    _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "First level counts of data mismatch. Range will be extended.");
                }

                if (dataCntSecondLevel > cntSecondLevel)
                {
                    _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Second level counts of data mismatch. Range will be extended.");
                }

                if (rowFirst)
                {
                    crange.EnsureCapcity(dataCntFirstLevel, dataCntSecondLevel);
                }
                else
                {
                    crange.EnsureCapcity(dataCntSecondLevel, dataCntFirstLevel);
                }
            }
            else
            {
                if (dataCntFirstLevel != cntFirstLevel)
                {
                    _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "First level counts of data mismatch. Data may be truncated.");
                }

                if (dataCntSecondLevel != cntSecondLevel)
                {
                    _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Second level counts of data mismatch. Data may be truncated.");
                }
            }

            var cellPositions = rowFirst ? crange.EnumerateRowFirst() : crange.EnumerateColumnFirst();

            using var dataEnumerator = data.Branches.GetEnumerator();

            foreach (var branch in cellPositions)
            {
                if (!dataEnumerator.MoveNext())
                    break;

                var curList = dataEnumerator.Current;
                var index = 0;

                foreach (var cref in branch)
                {
                    if (index >= curList.Count)
                        break;

                    var cellData = curList[index];

                    if (!(cellData is null && ignoreNull))
                    {
                        var cell = sheet.EnsureCell(cref.RowId, cref.ColumnId);
                        if (!CellAccessUtility.TrySetCellContent(cell, hint, cellData))
                        {
                            _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Cannot set the content of cell {cref}. Skipped.");
                        }
                    }

                    ++index;
                }
            }
        }
        public void ActualWriteData(
            ISheet sheet,
            List<SimpleCellReference> crefs,
            List<IGH_Goo> data,
            CellTypeHint hint,
            bool ignoreNull = false
            )
        {
            if (crefs.Count != data.Count)
                throw new InvalidOperationException("Cell reference and data must have same length.");

            for (var i = 0; i < crefs.Count; i++)
            {
                var cellData = data[i];
                var cref = crefs[i];

                if (cellData is null && ignoreNull)
                    continue;

                var cell = sheet.EnsureCell(cref.RowId, cref.ColumnId);
                if (!CellAccessUtility.TrySetCellContent(cell, hint, cellData))
                {
                    _docObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Cannot set the content of cell {cref}. Skipped.");
                }
            }
        }
    }
}
