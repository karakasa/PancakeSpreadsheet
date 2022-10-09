using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Grasshopper.Kernel.Types;
using Org.BouncyCastle.Utilities;
using PancakeSpreadsheet.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PancakeSpreadsheet.GH.Tweaks
{
    public class SetMutliCellUpgrader1 : IGH_UpgradeObject
    {
        public DateTime Version => new DateTime(2022, 10, 9, 0, 0, 0);

        private static readonly Guid FromId = new psSetMultiCellsOld().ComponentGuid;
        private static readonly Guid ToId = new psSetMultiCells().ComponentGuid;
        public Guid UpgradeFrom => FromId;

        public Guid UpgradeTo => ToId;

        public IGH_DocumentObject Upgrade(IGH_DocumentObject target, GH_Document document)
        {
            if (target is not IGH_Component comp)
                return null;

            if (GH_UpgradeUtil.SwapComponents(comp, UpgradeTo) is not psSetMultiCells newComp)
                return null;

            var param = new Param_Boolean()
            {
                Name = "Auto Extend",
                NickName = "E?",
                Description = psSetMultiCells.AutoExtendParamDesc,
                Access = GH_ParamAccess.item,
            };
            param.SetPersistentData(new GH_Boolean(false));
            newComp.Params.Input.Add(param);

            return newComp;
        }
    }
}
