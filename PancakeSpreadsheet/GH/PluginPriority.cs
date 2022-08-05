using Grasshopper;
using Grasshopper.Kernel;
using PancakeSpreadsheet.NpoiInterop;
using PancakeSpreadsheet.PancakeInterop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.GH
{
    public class PluginPriority : GH_AssemblyPriority
    {
        public override GH_LoadingInstruction PriorityLoad()
        {
            InteropServer.DetectPancakeAssembly();

            Instances.CanvasCreated += Instances_CanvasCreated;

            return GH_LoadingInstruction.Proceed;
        }

        private void Instances_CanvasCreated(Grasshopper.GUI.Canvas.GH_Canvas canvas)
        {
            InteropServer.DetectPancakeAssemblyAfterUiLoaded();

            Instances.ActiveCanvas.DocumentChanged += ActiveCanvas_DocumentChanged;
        }

        private void ActiveCanvas_DocumentChanged(Grasshopper.GUI.Canvas.GH_Canvas sender, Grasshopper.GUI.Canvas.GH_CanvasDocumentChangedEventArgs e)
        {
            if (e.OldDocument is not null)
            {
                e.OldDocument.SolutionStart -= CleanUpBeforeSolutionStart;
            }

            if (e.NewDocument is not null)
            {
                e.NewDocument.SolutionStart -= CleanUpBeforeSolutionStart;
                e.NewDocument.SolutionStart += CleanUpBeforeSolutionStart;
            }
        }

        private void CleanUpBeforeSolutionStart(object sender, GH_SolutionEventArgs e)
        {
            if (!GH_Document.EnableSolutions)
                return;

            // ObjectMonitor.CleanUp();
        }
    }
}
