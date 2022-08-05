using Grasshopper.Kernel.Types;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PancakeSpreadsheet.NpoiInterop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Params
{
    public class GooSpreadsheet : GH_Goo<WorkbookHolder>
    {
        public GooSpreadsheet()
        {

        }

        public GooSpreadsheet(WorkbookHolder value)
        {
            Value = value;
        }
        public override bool IsValid => Value != null;

        public override string TypeName => "Spreadsheet Workbook";

        public override string TypeDescription => "A virutal spreadsheet object containing one or more sheets.";

        public override IGH_Goo Duplicate()
        {
            throw new NotImplementedException();
        }

        public override string ToString()
        {
            return TypeName;
        }
    }
}
