using System;
using Microsoft.Win32;
using System.Windows;
using System.Globalization;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;


namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (3.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            else { 
                string[,] list;
                Excel.Application ObjWorkExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

                var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                int _columns = (int)lastCell.Column;
                int _rows = (int)lastCell.Row;
                list = new string[_rows, _columns];

                for (int j = 0; j < _columns; j++)
                    for (int i = 0; i < _rows; i++)
                        list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;

                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();

                using (ExcelLabEntities excelEntities = new ExcelLabEntities())
                {
                    for (int i = 1; i < 70; i++)
                    {
                        excelEntities.Users.Add(new User()
                        {
                            Client_ID = Convert.ToInt32(list[i, 0]),
                            Name = list[i, 1],
                            Date_Birth = list[i, 2],
                            //Date_Birth = DateTime.ParseExact(list[i, 3], "d", CultureInfo.InvariantCulture),
                            Index_ = list[i, 3],
                            City = list[i, 4],
                            Street = list[i, 5],
                            House = list[i, 6],
                            Flat = list[i, 7],
                            Email = list[i, 8]
                        });
                    }
                    excelEntities.SaveChanges();
                }
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
