using Grasshopper.Kernel.Parameters;
using PancakeSpreadsheet.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.NpoiInterop
{
    internal static class CellTypeHintNamedValues
    {
        private static void AddNamedValues(Param_Integer paramOption)
        {
            paramOption.AddNamedValue("Automatic", (int)CellTypeHint.Automatic);
            paramOption.AddNamedValue("Formula", (int)CellTypeHint.Formula);
            paramOption.AddNamedValue("Datetime", (int)CellTypeHint.Datetime);
            paramOption.AddNamedValue("Text", (int)CellTypeHint.Text);
        }

        public static void AddCellTypeHintValues(this PancakeComponent component)
        {
            var paramOption = component.Params.Input.Last() as Param_Integer;
            AddNamedValues(paramOption);
        }
    }
}
