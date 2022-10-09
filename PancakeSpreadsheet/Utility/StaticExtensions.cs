using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Utility
{
    internal static class StaticExtensions
    {
        public static R AsGoo<T, R>(this T val)
            where R : GH_Goo<T>, new()
        {
            return new R { Value = val };
        }

        public static void Swap<T>(ref T a, ref T b)
        {
            var c = a;
            a = b;
            b = c;
        }

        public static bool HasDuplicates<T>(this IEnumerable<T> src)
        {
            var hs = new HashSet<T>();
            foreach (var it in src)
            {
                if (it is null)
                    continue;
                if (!hs.Add(it))
                    return true;
            }    

            return false;
        }

        public static Stream OpenFile(string filepath)
        {
            if ((filepath.StartsWith(@"\\") || filepath.StartsWith(@"//")) && !filepath.StartsWith(@"\\?\"))
            {
                try
                {
                    // Use memory instead for network-located files

                    return CopyAndReadIntoMemoryStream(filepath);
                }
                catch
                {
                }
            }

            try
            {
                var fs = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                return fs;
            }
            catch
            {
                try
                {
                    var fs = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    return fs;
                }
                catch
                {
                    try
                    {
                        // Last Resort

                        return CopyAndReadIntoMemoryStream(filepath);
                    }
                    catch
                    {
                        throw new IOException("Failed to read the file.");
                    }
                }
            }
        }

        private static MemoryStream CopyAndReadIntoMemoryStream(string filepath)
        {
            var tempFile = Path.GetTempFileName();
            File.Copy(filepath, tempFile, true);

            var fileBytes = File.ReadAllBytes(tempFile);
            var memoryStream = new MemoryStream(fileBytes);

            File.Delete(tempFile);

            return memoryStream;
        }

        public static bool IsZeroPath(this GH_Path path)
        {
            return path.Indices.All(i => i == 0);
        }

        public static int FirstNonzeroPath(IList<GH_Path> paths)
        {
            for(var i = 0; i < paths.Count; i++)
            {
                var path = paths[i];
                if (!IsZeroPath(path))
                    return i;
            }

            return -1;
        }
    }
}
