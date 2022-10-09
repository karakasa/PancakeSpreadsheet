using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using PancakeSpreadsheet.Params;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public class psAppendLabels : PancakeComponent
    {
        public override Guid ComponentGuid => new("{316082A8-96E7-45FC-ACCF-66DCF7DDA1DD}");

        protected override string ComponentName => "Append Labels";

        protected override string ComponentDescription => "Append row (and/or) column labels onto a tree structure.";

        protected override string ComponentCategory => PancakeComponent.CategorySpreadsheet;
        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Data", "D", "Data to append", GH_ParamAccess.tree);
            pManager.AddBooleanParameter("Row first?", "R?", "Whether to format the data as row-first or column-first. By default row-first, that is, one branch per row.", GH_ParamAccess.item, true);

            pManager.AddTextParameter("Row Labels", "RL", "Row labels to append. Ignored if empty.", GH_ParamAccess.list);
            pManager.AddTextParameter("Column Labels", "CL", "Column labels to append. Ignored if empty.", GH_ParamAccess.list);

            Params.Input[2].Optional = true;
            Params.Input[3].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Data", "D", "Appended data.", GH_ParamAccess.tree);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            
        }
    }
}
