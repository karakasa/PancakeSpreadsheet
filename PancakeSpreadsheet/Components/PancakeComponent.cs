using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;

namespace PancakeSpreadsheet.Components
{
    public abstract class PancakeComponent : GH_Component
    {
        private FeaturePack _features;
        internal FeaturePack Features => _features ??= new(this);
        /// <summary>
        /// Initializes a new instance of the PancakeComponent class.
        /// </summary>
        public PancakeComponent()
          : base(string.Empty, string.Empty, string.Empty, string.Empty, string.Empty)
        {
            CopyFrom(new GH_InstanceDescription(ComponentName,
                ComponentNickname, ComponentDescription, "Spreadsheet", ComponentCategory));
        }
        protected virtual string ComponentNickname => GetType().Name;
        protected abstract string ComponentName { get; }
        protected abstract string ComponentDescription { get; }
        protected abstract string ComponentCategory { get; }
        //protected bool UnwrapAndCheck<T>(object input, out T value) where T : class
        //{
        //    if()
        //}

        internal const string CategorySpreadsheet = "01 | Spreadsheet";
        internal const string CategorySheet = "02 | Sheet";
        internal const string CategoryCellContent = "03 | Cell";
        internal const string CategoryCellStyle = "04 | Style";
    }
}