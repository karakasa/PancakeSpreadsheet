using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using PancakeSpreadsheet.NpoiInterop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Params
{
    public class ParamCellRangeReference : PancakeParam<GooCellRangeReference>
    {
        public override Guid ComponentGuid => new("{E21A5708-2DE6-44CE-847A-F883509F178F}");

        protected override string ComponentName => "Cell Range Reference Param";

        protected override string ComponentDescription => "Represents a reference to a range of cells.";
        public override GH_Exposure Exposure => GH_Exposure.hidden;

        protected override GooCellRangeReference PreferredCast(object data)
        {
            try
            {
                SimpleCellRange value;

                switch (data)
                {
                    case GH_String gh_str:
                        value = SimpleCellRange.FromString(gh_str.Value);
                        return new GooCellRangeReference { Value = value };
                    case string str:
                        value = SimpleCellRange.FromString(str);
                        return new GooCellRangeReference { Value = value };
                }
            }
            catch
            {
            }

            return null;
        }
    }
}
