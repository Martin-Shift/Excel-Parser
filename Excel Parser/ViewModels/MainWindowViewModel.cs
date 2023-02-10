using IronXL;
using Microsoft.Win32;
using My.BaseViewModels;
using System;
using System.Data;
using System.Windows.Controls;
using System.Windows;
using System.Windows.Input;
using Microsoft.Office.Interop.Excel;
using System.IO;
using Excel_Parser.Models;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Excel_Parser.ViewModels;

public class MainWindowViewModel : NotifyPropertyChangedBase
{
    public MainWindowViewModel()
    {

    }
    public ObservableCollection<ProductViewModel> Products
    {
        get
        {
            var collection = new ObservableCollection<ProductViewModel>();
            AllProducts.ForEach(p => collection.Add(new ProductViewModel(p)));
            return collection;
        }
    }
    private ProductViewModel _selectedProduct;
    public ProductViewModel SelectedProduct
    {
        get { return _selectedProduct; }
        set
        {
            _selectedProduct = value;
            OnPropertyChanged(nameof(SelectedProduct));
        }
    }


    public List<Product> AllProducts { get; set; } = new();

    private System.Data.DataTable _data { get; set; }
    public System.Data.DataTable Data { get => _data; set { _data = value; OnPropertyChanged(nameof(Data)); } }
    private WorkBook _Book { get; set; }
    public ICommand SaveFileAs => new RelayCommand(x =>
    {
        SaveFileDialog saveFile = new SaveFileDialog();
        saveFile.Filter = "Excel Document (*.xlsx)|*.xlsx";
        saveFile.DefaultExt = "xlsx";
        if (saveFile.ShowDialog() == true)
        {
            if (_Book != null)
            {
                var window = new SaveWindow(Data, saveFile.FileName);
                window.ShowDialog();
                MessageBox.Show("Saved!", "Save file", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Nothing to save!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    });
    public ICommand SaveFile => new RelayCommand(x =>
    {
        if (_Book != null)
        {
            File.Delete(_Book.FilePath);
            var window = new SaveWindow(Data, _Book.FilePath);
            window.ShowDialog();
            MessageBox.Show("Saved!", "Save file", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        else
        {
            MessageBox.Show("Nothing to save!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    });
    public ICommand OpenFromFile => new RelayCommand(async x =>
    {
        OpenFileDialog file = new();
        file.DefaultExt = "xlsx";
        file.Filter = "Excel Document (*.xlsx)|*.xlsx";
        if (file.ShowDialog() == true)
        {
            try
            {
                _Book = WorkBook.Load(file.FileName);
                WorkSheet sheet = _Book.DefaultWorkSheet;
                Data = sheet.ToDataTable(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
    }
    );
    public ICommand OpenFromFileProduct => new RelayCommand(x =>
    {
        OpenFileDialog file = new();
        file.DefaultExt = "xlsx";
        file.Filter = "Excel Document (*.xlsx)|*.xlsx";
        if (file.ShowDialog() == true)
        {
            AllProducts.Clear();
            _Book = WorkBook.Load(file.FileName);
            WorkSheet sheet = _Book.DefaultWorkSheet;
            var rowCount = sheet.RowCount;
            for (int i = 11; i < rowCount; i++)
            {
                var product = new Product
                {
                    Sku = sheet[$"B{i}"].StringValue,
                    ForReceiving = sheet[$"E{i}"].IntValue,
                    FreeStockQuantity = sheet[$"H{i}"].IntValue,
                    Order = sheet[$"G{i}"].IntValue,
                    Reserved = sheet[$"D{i}"].IntValue,
                    StockQuantity = sheet[$"C{i}"].IntValue,
                    Transfers = sheet[$"F{i}"].IntValue
                };
                if (product.Sku == null || product.Sku.Length == 0)
                {
                    break;
                }
                AllProducts.Add(product);
                OnPropertyChanged(nameof(Products));
            };
        }
    });
    public ICommand OpenFromExcelProduct => new RelayCommand(x =>
    {
        AllProducts.Clear();
        WorkSheet sheet = _Book.DefaultWorkSheet;
        var rowCount = sheet.RowCount;
        for (int i = 11; i < rowCount; i++)
        {
            var product = new Product
            {
                Sku = sheet[$"B{i}"].StringValue,
                ForReceiving = sheet[$"E{i}"].IntValue,
                FreeStockQuantity = sheet[$"H{i}"].IntValue,
                Order = sheet[$"G{i}"].IntValue,
                Reserved = sheet[$"D{i}"].IntValue,
                StockQuantity = sheet[$"C{i}"].IntValue,
                Transfers = sheet[$"F{i}"].IntValue
            };
            if (product.Sku == null || product.Sku.Length == 0)
            {
                break;
            }
            AllProducts.Add(product);
            OnPropertyChanged(nameof(Products));
        };

    });
    public ICommand SaveProducts => new RelayCommand(x =>
    {
        SaveFileDialog saveFile = new SaveFileDialog();
        saveFile.Filter = "Excel Document (*.xlsx)|*.xlsx";
        saveFile.DefaultExt = "xlsx";
        if (saveFile.ShowDialog() == true)
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");
            int row = 1;
            AllProducts.ForEach(x =>
            {
                workSheet[$"B{row}"].Value = x.Sku;
                workSheet[$"E{row}"].Value = x.ForReceiving;
                workSheet[$"H{row}"].Value = x.FreeStockQuantity;
                workSheet[$"G{row}"].Value = x.Order;
                workSheet[$"D{row}"].Value = x.Reserved;
                workSheet[$"C{row}"].Value = x.StockQuantity;
                workSheet[$"F{row}"].Value = x.Transfers;
                row++;
            });

            workBook.SaveAs(saveFile.FileName);
        }
    });

}
