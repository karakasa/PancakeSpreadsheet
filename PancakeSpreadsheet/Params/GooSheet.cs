using Grasshopper.Kernel.Types;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Params
{
    public class GooSheet : GH_Goo<ISheet>
    {
        public override bool IsValid => Value != null;

        public override string TypeName => "Sheet";

        public override string TypeDescription => "Represents a sheet inside a spreadsheet object.";

        public override IGH_Goo Duplicate()
        {
            throw new NotImplementedException();
        }

        public override string ToString()
        {
            return $"Sheet {Value?.SheetName ?? String.Empty}";
        }
    }
}
