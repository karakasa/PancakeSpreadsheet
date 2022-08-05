using Grasshopper.Kernel;
using PancakeSpreadsheet.PancakeInterop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public abstract class PancakeComponentRequireCore : PancakeComponent
    {
        protected GH_Exposure CalculateExposure(GH_Exposure desired)
            => ShouldBeVisible ? desired : GH_Exposure.hidden;
        protected bool ShouldBeVisible => InteropServer.IsPancakeAvailable;

        protected bool ShouldRunSolveInstance()
        {
            if (!InteropServer.IsPancakeAvailable)
            {
                AddRuntimeMessage(Grasshopper.Kernel.GH_RuntimeMessageLevel.Error, "This feature requires the core Pancake experience installed.");
                return false;
            }

            return true;
        }
    }
}
