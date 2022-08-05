using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;

namespace PancakeSpreadsheet.Params
{
    public abstract class PancakeParam<T> : GH_Param<T> where T : class, IGH_Goo
    {
        /// <summary>
        /// Initializes a new instance of the PancakeComponent class.
        /// </summary>
        public PancakeParam()
          : base(string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, GH_ParamAccess.item)
        {
            CopyFrom(new GH_InstanceDescription(ComponentName,
                GetType().Name, ComponentDescription, "Spreadsheet", "00 | Spreadsheet"));
        }

        protected abstract string ComponentName { get; }
        protected abstract string ComponentDescription { get; }

    }
}