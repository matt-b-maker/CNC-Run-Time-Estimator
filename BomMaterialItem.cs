using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPFSWTry
{
    public class BomMaterialItem
    {
        public BomMaterialItem()
        {
            Quantity = 0;
        }

        public BomMaterialItem(string partNo)
        {
            this.PartNo = partNo;
        }

        public BomMaterialItem(string description, string thickness)
        {
            this.Description = description;
            this.Thickness = thickness;
        }

        public string Description { get; set; }

        public string Thickness { get; set; }

        public string PartNo { get; set; }

        public int Quantity { get; set; }

        public double Price { get; set; }

        public double TotalPrice { get; set; }
    }
}
