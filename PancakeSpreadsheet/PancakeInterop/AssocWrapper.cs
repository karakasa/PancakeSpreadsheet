using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.PancakeInterop
{
    internal sealed class AssocWrapper
    {
        private readonly object _internalObj;
        public List<string> Names { get; private set; }
        public List<object> Values { get; private set; }
        public AssocWrapper()
        {
            _internalObj = AssocAccessor.CtorMethod();

            LoadList();
        }

        public AssocWrapper(object internalObj)
        {
            if (internalObj is null || internalObj.GetType() != AssocAccessor.AssocType)
            {
                throw new InvalidCastException("Target value isn't Pancake association.");
            }

            _internalObj = internalObj;

            LoadList();
        }

        public IGH_Goo AssocObject => _internalObj as IGH_Goo;

        private void LoadList()
        {
            Names = AssocAccessor.GetNamesMethod(_internalObj);
            Values = AssocAccessor.GetContentMethod(_internalObj);
        }

        public void Add(string name, object content)
        {
            Names.Add(PolishName(name));
            Values.Add(content);
        }

        public void Add(object content)
        {
            Add(null, content);
        }

        public object Get(string name)
        {
            for (var i = 0; i < Names.Count; i++)
                if (Names[i] == name)
                    return Values[i];

            return null;
        }

        public IEnumerator<KeyValuePair<string, object>> GetEnumerator()
        {
            for (var i = 0; i < Values.Count; i++)
            {
                yield return new KeyValuePair<string, object>(
                    (i >= Names.Count) ? null : Names[i],
                    Values[i]);
            }
        }
        private static string PolishName(string name)
        {
            if (name is null)
                return null;

            return name
                .Replace("\n", string.Empty)
                .Replace("\r", string.Empty)
                .Replace("\t", string.Empty)
                .Trim();
        }
    }
}
