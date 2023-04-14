using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.PancakeInterop
{
    internal static class InteropServer
    {
        private static readonly Lazy<bool> _shouldShowPancakeComponents = new(DetermineShouldShowPancakeComponents);
        public static bool ShouldShowPancakeComponents => _shouldShowPancakeComponents.Value;
        public static bool IsPancakeAvailable { get; private set; } = false;
        public static void DetectPancakeAssembly()
        {
            LocatePancakeInstallation();
        }

        private static bool DetermineShouldShowPancakeComponents()
        {
            const string PancakeFileName = "Pancake.gha";

            if (IsPancakeAvailable)
                return true;

            foreach (var assemblyFldr in Folders.AssemblyFolders)
            {
                var path = Path.Combine(assemblyFldr.Folder, PancakeFileName);
                if (File.Exists(path))
                    return true;
            }

            return false;
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
