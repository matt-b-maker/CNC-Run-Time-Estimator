using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPFSWTry
{
    public class BomHardwareItem : BomMaterialItem
    {
        public BomHardwareItem(string partNo)
        {
            this.PartNo = partNo;
            this.Thickness = "No thickness / Hardware item";
        }

        
    }
}
