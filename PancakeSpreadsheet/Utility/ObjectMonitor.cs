using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Utility
{
    public class ObjectMonitor
    {
        private List<IDisposable> _disposables = new();

        public void Add(IDisposable obj)
        {
            _disposables.Add(obj);
        }

        public void CleanUp()
        {
            foreach (var it in _disposables)
                it?.Dispose();

            _disposables.Clear();
        }
    }
}
