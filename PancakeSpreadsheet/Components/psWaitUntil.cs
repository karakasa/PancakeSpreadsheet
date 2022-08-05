using System;
using System.Drawing;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;

namespace PancakeSpreadsheet.Components
{
    public class psWaitUntil : PancakeComponent
    {
        protected override string ComponentCategory => PancakeComponent.CategorySpreadsheet;
        protected override string ComponentName => "Wait Until";
        protected override string ComponentDescription => "Postpone data until signals are received.\r\n" +
            "The component is required for some order-sensitive operations. For example, saving spreadsheet must be after all modifications.\r\n" +
            "The component is interchangable with the same-named component from the core Pancake experience.";
        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Signal", "S", "Signal", GH_ParamAccess.tree);
            pManager.AddGenericParameter("Data", "D", "Data to postpone", GH_ParamAccess.tree);
        }
        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Data", "D", "Data to postpone", GH_ParamAccess.tree);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="da">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess da)
        {
            da.GetDataTree<IGH_Goo>(1, out var tree);
            da.SetDataTree(0, tree);
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
            => ComponentIcons.WaitUntil;

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("2d334e87-a1ff-4a8e-85df-c0f5db0bee0a"); }
        }

        public override GH_Exposure Exposure => GH_Exposure.secondary;
    }
}