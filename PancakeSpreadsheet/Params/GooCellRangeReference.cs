using Grasshopper.Kernel.Types;
using PancakeSpreadsheet.NpoiInterop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Params
{
    public class GooCellReference : GH_Goo<SimpleCellReference>
    {
        public override bool IsValid => Value.IsValid();

        public override string TypeName => "Cell Reference";

        public override string TypeDescription => "Reference a cell position";

        public override IGH_Goo Duplicate()
        {
            return new GooCellReference { Value = Value };
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
                        Value = SimpleCellReference.FromString(gh_str.Value);
                        return true;
                    case string str:
                        Value = SimpleCellReference.FromString(str);
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
