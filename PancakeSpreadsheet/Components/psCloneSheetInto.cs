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
    public class psCloneSheetInto : PancakeComponent
    {
        protected override string ComponentCategory => PancakeComponent.CategorySheet;
        public override Guid ComponentGuid => new("{0C5A40AA-C443-4C1A-9932-0C2B18EE02A3}");

        protected override string ComponentName => "Clone Sheet Into";

        protected override string ComponentDescription => "Copy an existing sheet into another workbook.";

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Source spreadsheet", "SP(S)", "Source spreadsheet object", GH_ParamAccess.item);
            pManager.AddGenericParameter("Identifier", "I", "Name or index of the sheet", GH_ParamAccess.item);
            pManager.AddTextParameter("Name", "N", "New name of the cloned sheet. Name is kept by default.", GH_ParamAccess.item, string.Empty);

            pManager.AddGenericParameter("Target spreadsheet", "SP(T)", "Target spreadsheet object", GH_ParamAccess.item);

            pManager.AddBooleanParameter("Keep formula", "KF", "Keep formula. True by default.", GH_ParamAccess.item, true);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "The new sheet object", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooSpreadsheet goo = default;
            GooSpreadsheet goo2 = default;

            string newname = default;
            bool keepFormula = true;

            DA.GetData(2, ref newname);

            DA.GetData(0, ref goo);
            DA.GetData(3, ref goo2);
            DA.GetData(4, ref keepFormula);


            var wb = goo?.Value?.Workbook;
            var targetWb = goo2?.Value?.Workbook;

            if (wb is null || targetWb is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid spreadsheet.");
                return;
            }

            object idObj = default;

            DA.GetData(1, ref idObj);

            ISheet sheet = null;
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
                        sheet = wb.GetSheetAt(sheetId);
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

            if (sheet is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid sheet.");
                return;
            }

            if (string.IsNullOrEmpty(newname))
                newname = sheetName;

            sheet.CopyTo(targetWb, newname, true, keepFormula);

            DA.SetData(0, targetWb.GetSheet(newname).AsGoo());
        }

        protected override Bitmap Icon => ComponentIcons.CopySheetInto;
        public override GH_Exposure Exposure => GH_Exposure.secondary;
    }
}

