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
    public class psSimpleRead : PancakeComponent
    {
        public override Guid ComponentGuid => new("{286082A8-96E7-45FC-ACCF-66DCF7DDA1DD}");

        protected override string ComponentName => "Simple Read";

        protected override string ComponentDescription => "One unified component to read a simple xls(x) file.";

        protected override string ComponentCategory => PancakeComponent.CategorySpreadsheet;
        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddParameter(new Param_FilePath(), "File", "F", "File to read", GH_ParamAccess.item);
            pManager.AddGenericParameter("Sheet Identifier", "S", "Sheet name or index (0-based) to read. By default the first is read.", GH_ParamAccess.item);
            pManager.AddGenericParameter("Read Location", "RL", "Where to read. It can be a starting position, or a range, or nothing. By default the entire sheet is read.", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Row first?", "R?", "Whether to format the data as row-first or column-first. By default row-first, that is, one branch per row.", GH_ParamAccess.item, true);
            pManager.AddBooleanParameter("OK", "OK", "OK to read", GH_ParamAccess.item, false);

            Params.Input[1].Optional = true;
            Params.Input[2].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Data", "D", "Imported data.", GH_ParamAccess.tree);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            throw new NotImplementedException();
        }
    }
}
