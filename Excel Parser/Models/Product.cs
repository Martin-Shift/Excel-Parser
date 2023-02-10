using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Parser.Models
{
    public class Product
    {
        public string Sku { get; set; }
        public int StockQuantity { get; set; }
        public int Reserved { get; set; }
        public int Transfers { get; set; }
        public int ForReceiving { get; set; }
        public int Order { get; set; }
        public int FreeStockQuantity { get; set; }

    }
}
