
using Excel = Microsoft.Office.Interop.Excel;
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



namespace Template4335
{
    /// <summary>
    /// Логика взаимодействия для Nigametyanova4335.xaml
    /// </summary>
    public partial class Nigametyanova4335 : Window
    {
        public Nigametyanova4335()
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
            using (vtorEntities1 vtorEntities = new vtorEntities1())
            {
                for (int i = 1; i <= 50; i++)
                {
                    Uslugi uslugi = new Uslugi();
                    uslugi.Id = int.Parse(list[i, 0]);
                    uslugi.CodeZakaz = list[i, 1];
                    uslugi.Date = list[i, 2];
                    uslugi.TimeZakaz = list[i, 3];
                    uslugi.CodeClient = list[i, 4];
                    uslugi.Uslugi1 = list[i, 5];
                    uslugi.Status = list[i, 6];
                    uslugi.DateOff = list[i, 7];
                    uslugi.TimeProkat = list[i, 8];
                    vtorEntities.Uslugi.Add(uslugi);
                }
                vtorEntities.SaveChanges();
                MessageBox.Show("Импорт данных прошел успешно");
            }
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Uslugi> allUsl;
            List<string> allDate;
            using (vtorEntities1 vtorEntities = new vtorEntities1())
            {
                allUsl = vtorEntities.Uslugi.ToList();
                allDate = vtorEntities.Uslugi.ToList().Select(s => s.Date).Distinct().ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allDate.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < allDate.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(allDate[i]);
                worksheet.Cells[1][startRowIndex] = "ID";
                worksheet.Cells[2][startRowIndex] = "Код заказа";
                worksheet.Cells[3][startRowIndex] = "Код клиента";
                worksheet.Cells[4][startRowIndex] = "Услуги";
                
                startRowIndex++;
                var servicesCategories = allUsl.GroupBy(s => s.Date).ToList();
                foreach (var services in servicesCategories)
                {
                    if (services.Key == allDate[i])
                    {
                        foreach (Uslugi service in allUsl)
                        {
                            if (service.Date == services.Key)
                            {
                                worksheet.Cells[1][startRowIndex] = service.Id;
                                worksheet.Cells[2][startRowIndex] = service.CodeZakaz;
                                worksheet.Cells[3][startRowIndex] = service.CodeClient;
                                worksheet.Cells[4][startRowIndex] = service.Uslugi1;
              
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

       


    
