using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.Storage;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PancakeSpreadsheet.NpoiInterop;
using PancakeSpreadsheet.Params;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public class psSimpleWriteList : PancakeComponent
    {
        public override Guid ComponentGuid => new("{7F7652DF-0286-4F77-ADDF-F4F734D3D298}");

        protected override string ComponentName => "Simple Write List to XLS(X)";

        protected override string ComponentDescription => "One unified component to write a list of cells to a xls(x) file.";

        protected override string ComponentCategory => PancakeComponent.CategorySpreadsheet;
        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddParameter(new Param_FilePath(), "File", "F", "File to read", GH_ParamAccess.item);
            pManager.AddGenericParameter("Sheet Identifier", "S", "Sheet name or index (0-based) to write.", GH_ParamAccess.item);
            pManager.AddParameter(new ParamCellReference(), "Cell Positions", "CP", "Where to write data.", GH_ParamAccess.list);
            pManager.AddGenericParameter("Data", "D", "Data to write", GH_ParamAccess.list);
            pManager.AddBooleanParameter("Resize Column?", "RC?", "Resize all columns to fit content. By default true.", GH_ParamAccess.item, true);
            pManager.AddBooleanParameter("Create New?", "CN?", "If the output file exists, true to delete the file first, false to write onto the original file.", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("Ignore Null?", "IN?", "Nulls in the datatree are skipped when written to an existing file.", GH_ParamAccess.item, false);
            pManager.AddBooleanParameter("OK", "OK", "OK to write", GH_ParamAccess.item, false);

            Params.Input[1].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddBooleanParameter("OK", "OK", "Whether the operation is successful.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string filepath = default;
            IGH_Goo sheetId = default;
            var writePosition = new List<GooCellReference>();
            bool resizeCol = true;
            bool createNew = false;
            bool ignoreNull = false;
            bool ok = default;

            var dataList = new List<IGH_Goo>();

            DA.GetData(0, ref filepath);
            DA.GetData(1, ref sheetId);
            DA.GetDataList(2, writePosition);
            DA.GetDataList(3, dataList);
            DA.GetData(4, ref resizeCol);
            DA.GetData(5, ref createNew);
            DA.GetData(6, ref ignoreNull);
            DA.GetData(7, ref ok);

            if (!ok)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Wire a True value to the OK input, to write the file.");
                return;
            }

            if(dataList.Count == 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Nothing to write.");
                return;
            }

            if (writePosition.Count != dataList.Count)
            {
                if (dataList.Count == 1)
                {
                    dataList.Capacity = writePosition.Count;
                    var item = dataList[0];
                    for (var i = 1; i < writePosition.Count; i++)
                        dataList.Add(item);
                }
                else
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Data and position lists must have the same length.");
                    return;
                }
            }

            if (!Features.ValidateFile(filepath, out _, true))
                return;

            var holder = Features.OpenOrCreateFile(filepath, createNew, out var useExistingFile);

            if (holder is null)
                return;

            try
            {
                var sheet = Features.OpenOrCreateSheet(holder, sheetId, useExistingFile);
                if (sheet is null)
                    return;

                var positions = writePosition.Select(static goo => goo.Value).ToList();

                Features.ActualWriteData(sheet, positions, dataList, CellTypeHint.Automatic, ignoreNull);

                if (resizeCol)
                    Features.ResizeAll(sheet);

                Features.WriteToFile(filepath, holder, true);
            }
            finally
            {
                holder?.Dispose();
            }
        }
        protected override Bitmap Icon => ComponentIcons.SimpleWriteList;
    }
}
