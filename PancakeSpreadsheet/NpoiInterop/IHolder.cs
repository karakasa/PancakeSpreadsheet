using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.NpoiInterop
{
    public interface IHolder<T>
    {
        public T Value { get; }
    }
}
