using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using NPOI.SS.UserModel;
using PancakeSpreadsheet.NpoiInterop;
using PancakeSpreadsheet.Params;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public class psGetSheet : PancakeComponent
    {
        protected override string ComponentCategory => PancakeComponent.CategorySheet;
        public override Guid ComponentGuid => new("{0C5A40AA-C442-4C1A-9932-0C2B18EE02A4}");

        protected override string ComponentName => "Get Sheet";

        protected override string ComponentDescription => "Retrieve a sheet by name or index.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Spreadsheet", "SP", "Spreadsheet object", GH_ParamAccess.item);
            pManager.AddGenericParameter("Identifier", "I", "Name or index (0-based) of the sheet", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooSpreadsheet goo = default;

            DA.GetData(0, ref goo);

            var wb = goo?.Value?.Workbook;

            if (wb is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid spreadsheet.");
                return;
            }

            object idObj = default;

            DA.GetData(1, ref idObj);

            ISheet sheet;

            switch (idObj)
            {
                case GH_Number number:
                    try
                    {
                        sheet = wb.GetSheetAt((int)number.Value);
                    }
                    catch
                    {
                        sheet = null;
                    }

                    if (sheet is null)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Sheet index {number.Value} is invalid.");
                        return;
                    }

                    break;
                case GH_Integer index:
                    try
                    {
                        sheet = wb.GetSheetAt(index.Value);
                    }
                    catch
                    {
                        sheet = null;
                    }

                    if (sheet is null)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Sheet index {index.Value} is invalid.");
                        return;
                    }

                    break;
                case GH_String name:

                    sheet = wb.GetSheet(name.Value);

                    if (sheet is null)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Sheet {name.Value} doesn't exist.");
                        return;
                    }

                    break;
                default:
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Unknown identifier. Must be an integer or text");
                    return;
            }

            DA.SetData(0, sheet.AsGoo());
        }
        protected override Bitmap Icon => ComponentIcons.GetSheet;
    }
}
