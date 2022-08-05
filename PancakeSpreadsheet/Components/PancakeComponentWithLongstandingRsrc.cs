using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using PancakeSpreadsheet.Utility;
using Rhino.Geometry;

namespace PancakeSpreadsheet.Components
{
    public abstract class PancakeComponentWithLongstandingRsrc : PancakeComponent
    {
        private ObjectMonitor _resources = new();

        protected void MonitorResource(IDisposable resource)
            => _resources.Add(resource);
        protected override void BeforeSolveInstance()
        {
            // Clean up resources created previously
            _resources.CleanUp();

            base.BeforeSolveInstance();
        }

        public override void RemovedFromDocument(GH_Document document)
        {
            _resources.CleanUp();

            base.RemovedFromDocument(document);
        }
    }
}