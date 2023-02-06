using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using DataTable = System.Data.DataTable;

namespace Excel_Parser
{
    /// <summary>
    /// Interaction logic for SaveWindow.xaml
    /// </summary>
    public partial class SaveWindow : System.Windows.Window
    {
        public SaveWindow(System.Data.DataTable Data, string excelFilePath)
        {
            InitializeComponent();
            this.Data = Data;
            this.ExcelFilePath = excelFilePath; 
            SaveBar.Maximum = Data.Columns.Count * Data.Rows.Count + Data.Columns.Count;
        }
        public DataTable Data { get; set; }
        public string ExcelFilePath { get; set; }
        public void SaveExcel(System.Data.DataTable Data, string excelFilePath)
        {
            Close();
        }
        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {         
            if (Data == null || Data.Columns.Count == 0)
                throw new Exception("ExportToExcel: Null or empty input table!\n");
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Workbooks.Add();
            _Worksheet workSheet = excelApp.ActiveSheet;
            for (var i = 0; i < Data.Columns.Count; i++)
            {
                workSheet.Cells[1, i + 1] = Data.Columns[i].ColumnName;
            }
            for (var i = 0; i < Data.Rows.Count; i++)
            {
                for (var j = 0; j < Data.Columns.Count; j++)
                {
                    workSheet.Cells[i + 2, j + 1] = Data.Rows[i][j];
                }
            }
            try
            {
                workSheet.SaveAs(ExcelFilePath);
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"+ ex.Message);
            }
        }
        private void Window_ContentRendered(object sender, EventArgs e)
        {
            BackgroundWorker worker = new BackgroundWorker();
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            worker.DoWork += worker_DoWork;
            worker.RunWorkerAsync();
        }
        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Close();
        }
    }
}
