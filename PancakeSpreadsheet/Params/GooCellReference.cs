using Grasshopper.Kernel.Types;
using PancakeSpreadsheet.NpoiInterop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Params
{
    public class GooCellRangeReference : GH_Goo<SimpleCellRange>
    {
        public override bool IsValid => Value.IsValid();

        public override string TypeName => "Cell Range Reference";

        public override string TypeDescription => "Reference a rectangle of cells";

        public override IGH_Goo Duplicate()
        {
            return new GooCellRangeReference { Value = Value };
        }

        public override string ToString()
        {
            return Value.ToString();
        }
        public override bool CastFrom(object source)
        {
            try
            {
                switch (source)
                {
                    case GH_String gh_str:
                        Value = SimpleCellRange.FromString(gh_str.Value);
                        return true;
                    case string str:
                        Value = SimpleCellRange.FromString(str);
                        return true;
                }
            }
            catch
            {
                return false;
            }

            return base.CastFrom(source);
        }
    }
}
