using Grasshopper.Kernel.Types;
using NPOI.SS.UserModel;
using PancakeSpreadsheet.Params;
using System;
using System.IO;

namespace PancakeSpreadsheet.NpoiInterop
{
    public sealed class WorkbookHolder : IDisposable, IHolder<IWorkbook>
    {
        public IGH_Goo AsGoo() => new GooSpreadsheet(this);
        public Stream Stream { get; private set; }
        public IWorkbook Workbook { get; private set; }

        IWorkbook IHolder<IWorkbook>.Value => Workbook;

        private WorkbookHolder(Stream stream, IWorkbook wb)
        {
            Stream = stream;
            Workbook = wb;
        }

        public static WorkbookHolder Create(Stream stream, string password = "")
        {
            return Create(stream, WorkbookFactoryExtension.Create(stream, password));
        }
        public static WorkbookHolder Create(Stream stream, IWorkbook wb)
        {
            var holder = new WorkbookHolder(stream, wb);

            return holder;
        }

        public void Dispose()
        {
            Stream?.Dispose();
            Stream = null;

            Workbook?.Close();
            Workbook = null;

            GC.SuppressFinalize(this);
        }

        ~WorkbookHolder()
        {
            Dispose();
        }
        public bool IsFileBased => Stream is FileStream;
        public string BackendFile => (Stream as FileStream)?.Name;
    }
}
