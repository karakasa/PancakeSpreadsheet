using Grasshopper.Kernel;
using Grasshopper.Kernel.Special;
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
        public static IndexNameState TryGetIndexOrName(IGH_Goo goo, out int index, out string name)
        {
            switch (goo)
            {
                case GH_Integer ghInt:
                    index = ghInt.Value;
                    name = null;
                    return IndexNameState.Index;
                case GH_Number ghNumber:
                    try
                    {
                        index = (int)Math.Round(ghNumber.Value, MidpointRounding.AwayFromZero);
                        name = null;
                        return IndexNameState.Index;
                    }
                    catch
                    {
                    }
                    break;
                case GH_String ghString:
                    index = -1;
                    name = ghString.Value;
                    return IndexNameState.Name;
            }

            index = -1;
            name = null;
            return IndexNameState.Undefined;
        }
    }
    public enum IndexNameState
    {
        Undefined,
        Index,
        Name
    }
    
}
