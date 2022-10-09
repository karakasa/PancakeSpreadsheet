using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using PancakeSpreadsheet.Params;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.NpoiInterop
{
    internal enum CellTypeHint
    {
        Automatic = 0,
        Formula = 1,
        Datetime = 2,
        Text = 3
    }
    internal static class CellAccessUtility
    {
        public static ICell CellAt(this ISheet sheet, SimpleCellReference cref)
            => CellAt(sheet, cref.RowId, cref.ColumnId);
        public static ICell CellAt(this ISheet sheet, int rowId, int colId)
        {
            return sheet?.GetRow(rowId)?.GetCell(colId);
        }
        public static ICell EnsureCell(this ISheet sheet, SimpleCellReference cref)
            => EnsureCell(sheet, cref.RowId, cref.ColumnId);
        public static ICell EnsureCell(this ISheet sheet, int rowId, int colId)
        {
            var row = sheet.GetRow(rowId);
            if (row is null)
                row = sheet.CreateRow(rowId);

            var cell = row.GetCell(colId);
            if (cell is null)
                cell = row.CreateCell(colId);

            return cell;
        }
        public static CellTypeHint GetHint(int hintType)
        {
            if (hintType < 0 || hintType > 3) return CellTypeHint.Automatic;
            return (CellTypeHint)hintType;
        }

        public static bool IsEmpty(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Blank:
                    return true;
                case CellType.String:
                    return string.IsNullOrEmpty(cell.StringCellValue);
                default:
                    return false;
            }
        }
        public static bool TryGetCellContent(ICell cell, CellTypeHint hint, out object content)
        {
            switch (hint)
            {
                case CellTypeHint.Automatic:
                    // Automatic
                    return TryGetCellContentAutomatic(cell, out content);
                case CellTypeHint.Formula:
                    // Formula
                    if (cell.CellType == CellType.Formula)
                    {
                        content = cell.CellFormula;
                    }
                    else
                    {
                        content = string.Empty;
                        return false;
                    }
                    break;
                case CellTypeHint.Datetime:
                    // As datetime
                    content = cell.DateCellValue;
                    break;
                case CellTypeHint.Text:
                    // As text
                    content = cell.ToString();
                    break;
                default:
                    content = null;
                    return false;
            }

            return true;
        }

        public static bool TryGetCellContentAutomatic(ICell cell, out object content)
            => TryGetCellContent(cell, cell.CellType, out content);
        private static bool TryGetCellContent(ICell cell, CellType desiredType, out object content)
        {
            switch (desiredType)
            {
                case CellType.Numeric:
                    content = cell.NumericCellValue;
                    break;
                case CellType.String:
                    content = cell.StringCellValue;
                    break;
                case CellType.Formula:
                    return TryGetCellContent(cell, cell.CachedFormulaResultType, out content);
                    break;
                case CellType.Error:
                    content = FormulaError.ForInt(cell.ErrorCellValue).String;
                    break;
                case CellType.Blank:
                    content = string.Empty;
                    break;
                case CellType.Boolean:
                    content = cell.BooleanCellValue;
                    break;
                default:
                    content = null;
                    return false;
            }

            return true;
        }
        public static bool TrySetCellContent(ICell cell, CellTypeHint hint, IGH_Goo content)
        {
            switch (hint)
            {
                case CellTypeHint.Automatic:
                    // Automatic
                    return TrySetCellContent(cell, content);

                case CellTypeHint.Formula:
                    // Formula
                    cell.SetCellType(CellType.Formula);
                    cell.SetCellFormula((content?.ScriptVariable() ?? content)?.ToString() ?? String.Empty);
                    break;

                case CellTypeHint.Datetime:
                    // As datetime
                    if (!GH_Convert.ToDate(content, out var dt, GH_Conversion.Both))
                    {
                        return false;
                    }

                    cell.SetCellValue(dt);

                    break;
                case CellTypeHint.Text:
                    // As text

                    cell.SetCellType(CellType.String);
                    cell.SetCellValue((content?.ScriptVariable() ?? content)?.ToString() ?? String.Empty);
                    break;
            }

            return true;
        }
        public static bool TrySetCellContent(ICell cell, IGH_Goo content)
        {
            if (content is null)
            {
                cell.SetCellType(CellType.Blank);
                return true;
            }

            switch (content)
            {
                case GH_Integer gh_int:
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(gh_int.Value);
                    break;
                case GH_Number gh_double:
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(gh_double.Value);
                    break;
                case GH_Boolean gh_boolean:
                    cell.SetCellType(CellType.Boolean);
                    cell.SetCellValue(gh_boolean.Value);
                    break;
                case GH_String gh_str:
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(gh_str.Value);
                    break;
                default:
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(content.ToString());
                    return false;
            }

            return true;
        }

        public static bool TrySetCellContentPrimitive(ICell cell, CellTypeHint hint, object content)
        {
            switch (hint)
            {
                case CellTypeHint.Automatic:
                    // Automatic
                    return TrySetCellContentPrimitive(cell, content);

                case CellTypeHint.Formula:
                    // Formula
                    cell.SetCellType(CellType.Formula);
                    cell.SetCellFormula(content?.ToString() ?? String.Empty);
                    break;

                case CellTypeHint.Datetime:
                    // As datetime
                    if (!GH_Convert.ToDate(content, out var dt, GH_Conversion.Both))
                    {
                        return false;
                    }

                    cell.SetCellValue(dt);

                    break;
                case CellTypeHint.Text:
                    // As text

                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(content?.ToString() ?? String.Empty);
                    break;
            }

            return true;
        }
        public static bool TrySetCellContentPrimitive(ICell cell, object content)
        {
            if (content is null)
            {
                cell.SetCellType(CellType.Blank);
                return true;
            }

            switch (content)
            {
                case GH_Integer gh_int:
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(gh_int.Value);
                    break;
                case GH_Number gh_double:
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(gh_double.Value);
                    break;
                case GH_Boolean gh_boolean:
                    cell.SetCellType(CellType.Boolean);
                    cell.SetCellValue(gh_boolean.Value);
                    break;
                case GH_String gh_str:
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(gh_str.Value);
                    break;
                case int net_int:
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(net_int);
                    break;
                case double net_double:
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(net_double);
                    break;
                case bool net_boolean:
                    cell.SetCellType(CellType.Boolean);
                    cell.SetCellValue(net_boolean);
                    break;
                case string net_str:
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(net_str);
                    break;
                default:
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(content.ToString());
                    return false;
            }

            return true;
        }
        public static int GetRCIndex(IGH_Goo goo, bool isRow, out bool isArray, out int[] array)
        {
            isArray = false;
            array = null;

            switch (goo)
            {
                case GH_Integer gh_int:
                    return gh_int.Value;
                case GH_Number gh_number:
                    return (int)Math.Round(gh_number.Value);
                case GooCellReference cref:
                    if (!cref.Value.IsValid())
                        throw new ArgumentException(nameof(goo));
                    return isRow ? cref.Value.RowId : cref.Value.ColumnId;
                case GooCellRangeReference goo_crange:
                    var crange = goo_crange.Value;
                    if (!crange.IsValid())
                        throw new ArgumentException(nameof(goo));
                    array = CRangeToIndexes(crange, isRow).ToArray();
                    isArray = true;
                    return -1;
                case GH_String gh_str:
                    if (SimpleCellRange.TryFromString(gh_str.Value, out var crange2))
                    {
                        array = CRangeToIndexes(crange2, isRow).ToArray();
                        isArray = true;
                        return -1;
                    }
                    break;
            }

            throw new ArgumentException(nameof(goo));
        }

        private static IEnumerable<int> CRangeToIndexes(SimpleCellRange crange, bool isRow)
        {
            if (isRow)
            {
                return Enumerable.Range(crange.StartCell.RowId, crange.EndCell.RowId - crange.StartCell.RowId + 1);
            }
            else
            {
                return Enumerable.Range(crange.StartCell.ColumnId, crange.EndCell.ColumnId - crange.StartCell.ColumnId + 1);
            }
        }

        public enum CellDataType
        {
            Invalid,
            CellRef,
            CellRange
        }
        public static CellDataType TryGetCellData(IGH_Goo goo, out SimpleCellReference cref, out SimpleCellRange crange)
        {
            crange = default;
            cref = default;

            try
            {
                switch (goo)
                {
                    case GH_String ghStr:
                        var str = ghStr.Value;
                        if (SimpleCellReference.TryFromString(str, out cref))
                            return CellDataType.CellRef;
                        else if (SimpleCellRange.TryFromString(str, out crange))
                            return CellDataType.CellRange;
                        else
                            return CellDataType.Invalid;
                    case GooCellReference gooCell:
                        cref = gooCell.Value;
                        return CellDataType.CellRef;
                    case GooCellRangeReference gooCellRange:
                        crange = gooCellRange.Value;
                        return CellDataType.CellRange;
                    default:
                        return CellDataType.Invalid;
                }
            }
            catch
            {
                return CellDataType.Invalid;
            }
        }
    }
}
