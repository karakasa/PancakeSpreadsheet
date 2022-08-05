using Grasshopper.Kernel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.PancakeInterop
{
    internal static class InteropServer
    {
        public static bool IsPancakeAvailable { get; private set; } = false;
        public static void DetectPancakeAssembly()
        {
            LocatePancakeInstallation();
        }

        public static void DetectPancakeAssemblyAfterUiLoaded()
        {
            if (IsPancakeAvailable)
                return;

            // Pancake is loaded after this module

            LocatePancakeInstallation();
            if (IsPancakeAvailable)
            {
                // TODO: Notify components that require Pancake

                // Show hidden components

                GH_ComponentServer.UpdateRibbonUI();
            }
        }

        private static void LocatePancakeInstallation()
        {
            try
            {
                var assembly = AppDomain.CurrentDomain.GetAssemblies()
                .FirstOrDefault(a => (a?.GetName()?.Name?.Equals("Pancake") ?? false));

                if (assembly is null)
                    return;

                var assocType = assembly.GetType("Pancake.GH.Params.GhAssoc");
                if (assocType is null)
                    return;

                AssocAccessor.BindToType(assocType);

                IsPancakeAvailable = true;
            }
            catch (Exception ex)
            {
                // Slient fail; something is wrong.
            }
        }
    }
}
