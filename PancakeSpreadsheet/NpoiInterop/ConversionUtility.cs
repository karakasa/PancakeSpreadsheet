using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using NPOI.SS.UserModel;
using PancakeSpreadsheet.Params;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.NpoiInterop
{
    internal static class ConversionUtility
    {
        public static IGH_Goo AsGoo(this ISheet sheet)
        {
            return new GooSheet { Value = sheet };
        }

        public static object AsPrimitive(this IGH_Goo goo)
        {
            return goo.ScriptVariable();
        }
        public static bool TryGetIndex(IGH_Goo goo, out int index)
        {
            switch (goo)
            {
                case GH_Integer ghInt:
                    index = ghInt.Value;
                    return true;
                case GH_Number ghNumber:
                    try
                    {
                        index = (int)Math.Round(ghNumber.Value, MidpointRounding.AwayFromZero);
                        return true;
                    }
                    catch
                    {
                        index = -1;
                        return false;
                    }
            }

            index = -1;
            return false;
        }
    }
}
