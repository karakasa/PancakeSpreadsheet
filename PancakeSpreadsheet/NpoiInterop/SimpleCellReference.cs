using Grasshopper.Kernel.Types;
using NPOI.SS.Util;
using PancakeSpreadsheet.Params;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.NpoiInterop
{
    public struct SimpleCellReference
    {
        public int RowId;
        public int ColumnId;

        public static readonly SimpleCellReference Unset = new(-1, -1);
        public SimpleCellReference(int row, int col)
        {
            RowId = row;
            ColumnId = col;
        }

        public static SimpleCellReference FromString(string str)
        {
            if (str.Contains("R") && str.Contains("C"))
            {
                if (TryResolveR1C1Notation(str.Trim(), out var cref2))
                    return cref2;
            }

            var cref = new CellReference(str.Trim());
            return new(cref.Row, cref.Col);
        }

        public static bool TryFromString(string str, out SimpleCellReference cref)
        {
            try
            {
                cref = FromString(str);
                return true;
            }
            catch
            {
                cref = default;
                return false;
            }
        }

        public override string ToString()
        {
            return $"R{RowId + 1}C{ColumnId + 1}";
        }

        public IGH_Goo AsGoo()
            => new GooCellReference { Value = this };

        public CellReference AsNpoiObj()
            => new(RowId, ColumnId);

        public bool IsValid()
            => (RowId >= 0 && RowId < 16384) && (ColumnId >= 0 && ColumnId < 1048576);
        private static bool TryResolveR1C1Notation(string notation, out SimpleCellReference cellRef)
        {
            if (notation.Contains("["))
            {
                throw new NotSupportedException("Relative R1C1 notation is not supported.");
            }

            var indexR = notation.IndexOf("R");
            var indexC = notation.IndexOf("C");

            if (indexR >= indexC - 1)
            {
                cellRef = default;
                return false;
            }

            var strR = notation.Substring(indexR + 1, indexC - indexR - 1);
            var strC = notation.Substring(indexC + 1);

            if (!int.TryParse(strR, out var idR) || !int.TryParse(strC, out var idC)
                || idR <= 0 || idC <= 0)
            {
                cellRef = default;
                return false;
            }

            cellRef = new(idR - 1, idC - 1);
            return true;
        }
    }
}
