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
    public class psSimpleWrite : PancakeComponent
    {
        public override Guid ComponentGuid => new("{7F7652DF-0286-4F77-ADDF-F4F734D3D297}");

        protected override string ComponentName => "Simple Write";

        protected override string ComponentDescription => "One unified component to write a simple xls(x) file.";

        protected override string ComponentCategory => PancakeComponent.CategorySpreadsheet;
        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddParameter(new Param_FilePath(), "File", "F", "File to read", GH_ParamAccess.item);
            pManager.AddGenericParameter("Sheet Identifier", "S", "Sheet name or index (0-based) to write.", GH_ParamAccess.item);
            pManager.AddParameter(new ParamCellReference(), "Write Position", "WP", "Where to start writing. By default the top-left corner is used.", GH_ParamAccess.item);
            pManager.AddGenericParameter("Data", "D", "Data to write", GH_ParamAccess.tree);
            pManager.AddBooleanParameter("Row first?", "R?", "Whether to format the data as row-first or column-first. By default row-first, that is, one branch per row.", GH_ParamAccess.item, true);
            pManager.AddBooleanParameter("Resize Column?", "RC?", "Resize all columns to fit content. By default true.", GH_ParamAccess.item, true);
            pManager.AddBooleanParameter("Create New?", "CN?", "If the output file exists, true to delete the file first, false to write onto the original file.", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("OK", "OK", "OK to write", GH_ParamAccess.item, false);

            Params.Input[1].Optional = true;
            Params.Input[2].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddBooleanParameter("OK", "OK", "Whether the operation is successful.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            throw new NotImplementedException();
        }
    }
}
