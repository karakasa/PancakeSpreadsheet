using System;
using System.Linq;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Rhino.Geometry;
using PancakeSpreadsheet.Params;

using NPOI.SS.UserModel;
using System.IO;
using PancakeSpreadsheet.Utility;
using PancakeSpreadsheet.NpoiInterop;

namespace PancakeSpreadsheet.Components
{
    public class psOpenSpreadsheet : PancakeComponentWithLongstandingRsrc
    {
        protected override string ComponentCategory => PancakeComponent.CategorySpreadsheet;
        protected override string ComponentName => "Open Spreadsheet";
        protected override string ComponentDescription => "Open a spreadsheet.\r\n" +
            "Please be advised that password support is incomplete.";

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddParameter(new Param_FilePath(), "File", "F", "File to open", GH_ParamAccess.item);
            pManager.AddTextParameter("Password", "P", "Password of the file. Empty for no password.", GH_ParamAccess.item, string.Empty);
            pManager.AddIntegerParameter("Open Mode", "M", "Mode of opening. Right-click for more options. Default is suitable for most cases.\r\nContent-only ignores non-text info (pictures, diagrams, etc.), while the file cannot be saved back as original.", GH_ParamAccess.item, 0);

            var paramOpenMode = Params.Input.Last() as Param_Integer;
            paramOpenMode.AddNamedValue("Default", 0);
            paramOpenMode.AddNamedValue("Content only", 1);

            pManager.AddBooleanParameter("OK", "OK", "Whether to proceed\r\nBecause loading a large spreadsheet can be slow, you may connect a true-only button to control behaviors.", GH_ParamAccess.item, true);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Spreadsheet", "SP", "Spreadsheet object", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string filepath = default;
            string password = default;
            int openMode = default;
            bool ok = default;

            DA.GetData(0, ref filepath);
            DA.GetData(1, ref password);
            DA.GetData(2, ref openMode);
            DA.GetData(3, ref ok);

            if (!ok)
                return;

            if (openMode == 1)
            {
                WorkbookFactory.SetImportOption(ImportOption.SheetContentOnly);
            }
            else
            {
                WorkbookFactory.SetImportOption(ImportOption.All);
            }

            if (!File.Exists(filepath))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "File not found.");
                return;
            }

            var extension = Path.GetExtension(filepath).ToLowerInvariant();
            switch (extension)
            {
                case ".csv":
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "CSV file is unsupported. Use Pancake for CSV import & export.");
                    return;
                case ".xlsb":
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "XLSB file is unsupported.");
                    return;
            }

            var fileLength = new FileInfo(filepath).Length;
            if (fileLength < 32)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Empty file.");
                return;
            }

            const long InPlaceFileSizeThreshold = 1L * 1024 * 1024 * 1024; // 1GB

            Stream stream;

            if (fileLength > InPlaceFileSizeThreshold)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "File is too large. Pancake will open the file in-place. This may cause issues.");
                stream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }
            else
            {
                stream = new MemoryStream((int)fileLength);
                using var fileStream = StaticExtensions.OpenFile(filepath);

                fileStream.CopyTo(stream);
            }

            if (!string.IsNullOrEmpty(password))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Be advised that password support is incomplete. You may run into issues.");
            }

            try
            {
                var holder = WorkbookHolder.Create(stream, password);

                MonitorResource(holder);

                DA.SetData(0, holder.AsGoo());
            }
            catch (NotSupportedException nsEx)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, nsEx.Message);
                return;
            }
            catch (IOException ioEx)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, ioEx.InnerException?.Message ?? ioEx.Message);
                return;
            }
            
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return ComponentIcons.OpenFile;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("9B84B985-EF3B-484C-8CA4-923736C89EB7"); }
        }
    }
}