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

namespace Excel_Parser.ViewModels;

public class MainWindowViewModel : NotifyPropertyChangedBase
{
    public MainWindowViewModel()
    {

    }

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
}
