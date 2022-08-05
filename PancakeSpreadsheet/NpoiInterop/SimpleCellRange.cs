using Grasshopper.Kernel.Types;
using NPOI.SS.Util;
using PancakeSpreadsheet.Params;
using PancakeSpreadsheet.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.NpoiInterop
{
    public struct SimpleCellRange
    {
        public SimpleCellReference StartCell;
        public SimpleCellReference EndCell;

        public int ColumnCount => EndCell.ColumnId - StartCell.ColumnId + 1;
        public int RowCount => EndCell.RowId - StartCell.RowId + 1;
        public bool IsValid()
            => StartCell.IsValid() && EndCell.IsValid()
            && StartCell.RowId <= EndCell.RowId && StartCell.ColumnId <= EndCell.ColumnId;
        public SimpleCellRange(SimpleCellReference start, SimpleCellReference end)
        {
            StartCell = start;
            EndCell = end;
        }
        public override string ToString()
        {
            return $"{StartCell}:{EndCell}";
        }
        public IGH_Goo AsGoo()
        {
            return new GooCellRangeReference { Value = this };
        }
        public static bool TryFromString(string str, out SimpleCellRange crange)
        {
            var index = str.IndexOf(':');

            if (index == -1)
            {
                if (SimpleCellReference.TryFromString(str, out var cref))
                {
                    crange = new(cref, cref);
                    return true;
                }
                else
                {
                    crange = default;
                    return false;
                }
            }
            else
            {
                var firstCellStr = str.Substring(0, index);
                var lastCellStr = str.Substring(index + 1);

                if (SimpleCellReference.TryFromString(firstCellStr, out var startCell)
                    && SimpleCellReference.TryFromString(lastCellStr, out var endCell))
                {
                    crange = new(startCell, endCell);
                    return true;
                }
                else
                {
                    crange = default;
                    return false;
                }
            }
        }
        public static SimpleCellRange FromNpoiObj(CellRangeAddress crange)
        {
            return new SimpleCellRange
            {
                StartCell = new(crange.FirstRow, crange.FirstColumn),
                EndCell = new(crange.LastRow, crange.LastColumn)
            };
        }
        public static SimpleCellRange FromString(string str)
        {
            if (!TryFromString(str, out var crange))
                throw new InvalidCastException($"{str} cannot be casted to a cell range.");

            return crange;
        }
        public CellRangeAddress AsNpoiObj() 
            => new(StartCell.RowId, EndCell.RowId, StartCell.ColumnId, EndCell.ColumnId);

        public IEnumerable<IEnumerable<SimpleCellReference>> EnumerateRowFirst()
        {
            var rangeIterator = Enumerable.Range(StartCell.ColumnId, EndCell.ColumnId - StartCell.ColumnId + 1);

            for (var rowId = StartCell.RowId; rowId <= EndCell.RowId; rowId++)
                yield return rangeIterator.Select(colId => new SimpleCellReference(rowId, colId));
        }
        public IEnumerable<IEnumerable<SimpleCellReference>> EnumerateColumnFirst()
        {
            var rangeIterator = Enumerable.Range(StartCell.RowId, EndCell.RowId - StartCell.RowId + 1);

            for (var colId = StartCell.ColumnId; colId <= EndCell.ColumnId; colId++)
                yield return rangeIterator.Select(rowId => new SimpleCellReference(rowId, colId));
        }
        public IEnumerable<SimpleCellReference> Enumerate()
        {
            for (var rowId = StartCell.RowId; rowId <= EndCell.RowId; rowId++)
                for (var colId = StartCell.ColumnId; colId <= EndCell.ColumnId; colId++)
                    yield return new SimpleCellReference(rowId, colId);
        }
    }
}
