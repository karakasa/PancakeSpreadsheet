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
    public class ParamCellReference : PancakeParam<GooCellReference>
    {
        public override Guid ComponentGuid => new("{E21A5708-2DE6-44CE-847A-F883509F187F}");

        protected override string ComponentName => "Cell Reference Param";

        protected override string ComponentDescription => "Represents a reference to cell.";
        public override GH_Exposure Exposure => GH_Exposure.hidden;

        protected override GooCellReference PreferredCast(object data)
        {
            try
            {
                SimpleCellReference value;

                switch (data)
                {
                    case GH_String gh_str:
                        value = SimpleCellReference.FromString(gh_str.Value);
                        return new GooCellReference { Value = value };
                    case string str:
                        value = SimpleCellReference.FromString(str);
                        return new GooCellReference { Value = value };
                }
            }
            catch
            {
            }

            return null;
        }
    }
}
