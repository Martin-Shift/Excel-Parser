using My.BaseViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel_Parser.Models;

namespace Excel_Parser.ViewModels
{
    public class ProductViewModel : NotifyPropertyChangedBase
    {
        public Product Product { get; set; }
        public ProductViewModel(Product product)
        {
            Product = product;
        }
        public string Sku
        {
            get => Product.Sku; 
            set
            {
                Product.Sku = value;
                OnPropertyChanged(nameof(Sku));
            }
        }
        public int StockQuantity
        {
            get => Product.StockQuantity;
            set
            {
                Product.StockQuantity = value;
                OnPropertyChanged(nameof(StockQuantity));
            }
        }
        public int Reserved
        {
            get => Product.Reserved;
            set
            {
                Product.Reserved = value;
                OnPropertyChanged(nameof(Reserved));
            }
        }
        public int Transfers
        {
            get => Product.Transfers;
            set
            {
                Product.Reserved = value;
                OnPropertyChanged(nameof(Transfers));
            }
        }
        public int ForReceiving
        {
            get => Product.ForReceiving;
            set
            {
                Product.ForReceiving = value;
                OnPropertyChanged(nameof(ForReceiving));
            }
        }
        public int Order
        {
            get => Product.Order;
            set
            {
                Product.Order = value;
                OnPropertyChanged(nameof(Order));
            }
        }
        public int FreeStockQuantity
        {
            get => Product.FreeStockQuantity;
            set
            {
                Product.FreeStockQuantity = value;
                OnPropertyChanged(nameof(FreeStockQuantity));
            }
        }
    }
}
