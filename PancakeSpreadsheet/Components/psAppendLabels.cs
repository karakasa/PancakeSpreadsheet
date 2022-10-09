using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using PancakeSpreadsheet.Params;
using PancakeSpreadsheet.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.Components
{
    public class psAppendLabels : PancakeComponent
    {
        public override Guid ComponentGuid => new("{316082A8-96E7-45FC-ACCF-66DCF7DDA1DD}");

        protected override string ComponentName => "Append Labels";

        protected override string ComponentDescription => "Append row (and/or) column labels onto a tree structure.";

        protected override string ComponentCategory => PancakeComponent.CategorySpreadsheet;
        public override GH_Exposure Exposure => GH_Exposure.tertiary;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddGenericParameter("Data", "D", "Data to append", GH_ParamAccess.tree);
            pManager.AddBooleanParameter("Row first?", "R?", "Whether to format the data as row-first or column-first. By default row-first, that is, one branch per row.", GH_ParamAccess.item, true);

            pManager.AddTextParameter("Row Labels", "RL", "Row labels to append. Ignored if empty.", GH_ParamAccess.list);
            pManager.AddTextParameter("Column Labels", "CL", "Column labels to append. Ignored if empty.", GH_ParamAccess.list);

            Params.Input[2].Optional = true;
            Params.Input[3].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Data", "D", "Appended data.", GH_ParamAccess.tree);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool rowFirst = true;
            var listRowLabels = new List<string>();
            var listColumnLabels = new List<string>();

            DA.GetDataTree<IGH_Goo>(0, out var tree);
            DA.GetData(1, ref rowFirst);
            DA.GetDataList(2, listRowLabels);
            DA.GetDataList(3, listColumnLabels);

            var hasRowLabel = listRowLabels.Count > 0;
            var hasColLabel = listColumnLabels.Count > 0;

            if (!hasRowLabel && !hasColLabel)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "No labels supplied.");
                DA.SetDataTree(0, tree);
                return;
            }

            if (tree.PathCount == 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Input data is empty.");
                DA.SetDataTree(0, tree);
                return;
            }

            tree = tree.Duplicate();

            if (!rowFirst)
            {
                StaticExtensions.Swap(ref listRowLabels, ref listColumnLabels);
            }

            if (hasRowLabel)
            {
                var rowLabels = listRowLabels.Select(row => new GH_String(row)).ToArray();
                var singleZeroPath = new GH_Path(0);

                var paths = tree.Paths;
                if (tree.PathExists(singleZeroPath))
                {
                    var pathIds = paths
                        .Where(path => path.Indices.Length == 1)
                        .Select(path => path.Indices[0])
                        .OrderByDescending(i => i);

                    foreach (var i in pathIds)
                    {
                        tree.ReplacePath(new GH_Path(i), new GH_Path(i + 1));
                    }
                }

                tree.AppendRange(rowLabels, singleZeroPath);
            }

            if (hasColLabel)
            {
                if (hasRowLabel)
                    listColumnLabels.Insert(0, string.Empty);

                var index = 0;

                foreach (var list in tree.Branches)
                {
                    if (index >= listColumnLabels.Count)
                    {
                        list.Insert(0, new GH_String(string.Empty));
                    }
                    else
                    {
                        list.Insert(0, new GH_String(listColumnLabels[index]));
                    }

                    ++index;
                }
            }

            DA.SetDataTree(0, tree);
        }
    }
}
