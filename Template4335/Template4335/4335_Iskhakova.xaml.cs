using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4335
{
    /// <summary>
    /// Логика взаимодействия для _4335_Iskhakova.xaml
    /// </summary>
    public partial class _4335_Iskhakova : Window
    {
        public _4335_Iskhakova()
        {
            InitializeComponent();
        }

        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (1.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
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
            using (LR3Entities lrEntities = new LR3Entities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    Services services = new Services();
                    services.ID = int.Parse(list[i, 0]);
                    services.Name = list[i, 1];
                    services.Type = list[i, 2];
                    services.Code = list[i, 3];
                    services.Cost = int.Parse(list[i, 4]);
                    lrEntities.Services.Add(services);
                }
                lrEntities.SaveChanges();
                MessageBox.Show("Импорт данных прошел успешно");
            }
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Services> allServices;
            List<string> allType;
            using (LR3Entities lrEntities = new LR3Entities())
            {
                allServices = lrEntities.Services.ToList().OrderBy(s => s.Cost).ToList();
                allType = lrEntities.Services.ToList().Select(s => s.Type).Distinct().ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allType.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < allType.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(allType[i]);
                worksheet.Cells[1][startRowIndex] = "ID";
                worksheet.Cells[2][startRowIndex] = "Название услуги";
                worksheet.Cells[3][startRowIndex] = "Стоимость";
                startRowIndex++;
                var servicesCategories = allServices.GroupBy(s => s.Type).ToList();
                foreach (var services in servicesCategories)
                {
                    if (services.Key == allType[i])
                    {
                        foreach (Services service in allServices)
                        {
                            if (service.Type == services.Key)
                            {
                                worksheet.Cells[1][startRowIndex] = service.ID;
                                worksheet.Cells[2][startRowIndex] = service.Name;
                                worksheet.Cells[3][startRowIndex] = service.Cost;
                                startRowIndex++;
                            }
                        }
                    }
                    else
                    {
                        continue;
                    }
                }
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }
    }
}
