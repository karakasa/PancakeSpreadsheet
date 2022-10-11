﻿using Grasshopper.Kernel;
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
    public class psResizeAllColumns : PancakeComponent
    {
        public override GH_Exposure Exposure => GH_Exposure.tertiary;
        public override Guid ComponentGuid => new("{E0FF0CF3-0A2D-4924-B5E3-179DF57DF1D0}");

        protected override string ComponentName => "Resize All Columns";

        protected override string ComponentDescription => "Fit all columns with content.";

        protected override string ComponentCategory => CategoryCellStyle;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Sheet", "S", "Sheet object", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            GooSheet gooSheet = default;

            DA.GetData(0, ref gooSheet);

            var sheet = gooSheet?.Value;

            if (sheet is null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid sheet.");
                return;
            }

            ResizeAll(sheet);

            DA.SetData(0, gooSheet);
        }

        internal static void ResizeAll(ISheet sheet)
        {
            var firstRowIndex = sheet.FirstRowNum;
            var lastRowIndex = sheet.LastRowNum;

            var firstColIndex = int.MaxValue;
            var lastColIndex = int.MinValue;

            for (var rowIndex = firstRowIndex; rowIndex <= lastRowIndex; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row is null)
                    continue;

                if (row.FirstCellNum < firstColIndex)
                    firstColIndex = row.FirstCellNum;

                if (row.LastCellNum - 1 > lastColIndex)
                    lastColIndex = row.LastCellNum - 1;
            }

            for (var colIndex = firstColIndex; colIndex <= lastColIndex; colIndex++)
            {
                try
                {
                    sheet.AutoSizeColumn(colIndex);
                }
                catch
                {
                }
            }
        }

        // protected override Bitmap Icon => ComponentIcons.AdjustColumnWidth;
    }
}