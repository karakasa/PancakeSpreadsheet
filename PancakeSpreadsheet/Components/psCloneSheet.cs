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
    public class psCloneSheet : PancakeComponent
    {
        protected override string ComponentCategory => PancakeComponent.CategorySheet;
        public override Guid ComponentGuid => new("{0C5A40AA-C443-4C1A-9932-0C2B18EE02A4}");

        protected override string ComponentName => "Clone Sheet";

        protected override string ComponentDescription => "Create a sheet by cloning an existing one.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Spreadsheet", "SP", "Spreadsheet object", GH_ParamAccess.item);
            pManager.AddGenericParameter("Identifier", "I", "Name or index of the sheet", GH_ParamAccess.item);
            pManager.AddTextParameter("Name", "N", "New name of the cloned sheet", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "The new sheet object", GH_ParamAccess.item);
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

            string newname = default;

            DA.GetData(2, ref newname);

            ISheet sheet;
            int sheetId;
            string sheetName;

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

                    sheetId = (int)number.Value;
                    sheetName = sheet.SheetName;

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

                    sheetId = index.Value;
                    sheetName = sheet.SheetName;

                    break;
                case GH_String name:

                    try
                    {
                        sheetId = wb.GetSheetIndex(name.Value);
                    }
                    catch
                    {
                        sheetId = -1;
                    }

                    if (sheetId < 0)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Sheet {name.Value} doesn't exist.");
                        return;
                    }

                    sheetName = name.Value;

                    break;
                default:
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Unknown identifier. Must be an integer or text");
                    return;
            }

            if (sheetName == newname || string.IsNullOrEmpty(newname))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "New name is invalid.");
                return;
            }

            var newSheet = wb.CloneSheet(sheetId);
            wb.SetSheetName(wb.GetSheetIndex(newSheet), newname);

            DA.SetData(0, newSheet.AsGoo());
        }

        protected override Bitmap Icon => ComponentIcons.CloneSheet;
        public override GH_Exposure Exposure => GH_Exposure.secondary;
    }
}
