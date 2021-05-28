using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPFSWTry
{
    class OfficialBomItem : BomMaterialItem
    {
        public OfficialBomItem(string partNo, int quantity)
        {
            this.PartNo = partNo;
            this.Quantity = quantity;
            this.Description = "n/a";
            this.Thickness = "Hardware item";
            this.Price = 0.00;
            this.TotalPrice = 0.00;
        }

        public OfficialBomItem(string description, string thickness, int quantity)
        {
            this.Description = description;
            this.Thickness = thickness;
            this.Quantity = quantity;
            this.Price = 0.00;
            this.PartNo = "n/a";
            this.TotalPrice = 0.00;
        }
    }
}
