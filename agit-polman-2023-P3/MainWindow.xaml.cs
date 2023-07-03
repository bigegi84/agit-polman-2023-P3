using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace agit_polman_2023_P3
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private DataTableCollection tableCollection;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" };
            {
                if (openFileDialog.ShowDialog() == true)
                {
                    txtFilename.Text = openFileDialog.FileName;
                    //using var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read);
                    //using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                    //DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    //{
                    //    ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    //});
                    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        // Auto-detect format, supports:
                        //  - Binary Excel files (2.0-2003 format; *.xls)
                        //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            // Choose one of either 1 or 2:

                            // 1. Use the reader methods
                            do
                            {
                                while (reader.Read())
                                {
                                    // reader.GetDouble(0);
                                }
                            } while (reader.NextResult());

                            // 2. Use the AsDataSet extension method
                            var result = reader.AsDataSet();

                            // The result of each spreadsheet is in result.Tables
                            tableCollection = result.Tables;
                            cboSheet.Items.Clear();
                            foreach (DataTable table in tableCollection)
                                cboSheet.Items.Add(table.TableName);//Add  sheet to combobox
                        }
                    }
                    
                }
            }

        }

        private void cboSheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DataTable dt = tableCollection[cboSheet.SelectedItem.ToString()];
            dataGridView1.ItemsSource = dt.DefaultView;
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.CanUserAddRows = false;
        }
        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //DataTable dt = tableCollection[cboSheet.SelectedItem.ToString()];
            //dataGridView1.ItemsSource = dt.DefaultView;
            //dataGridView1.AutoGenerateColumns = true;
            //dataGridView1.CanUserAddRows = false;
        }
    }
}
