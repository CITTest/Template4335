using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4335
{

    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Khasanshin4335(object sender, RoutedEventArgs e)
        {
            Khasanshin4335 khas = new Khasanshin4335();
            khas.Show();
            Close();
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "Файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
            {
                return;
            }

            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (importisrpo2Entities1 usersEntities = new importisrpo2Entities1())
            {
                for (int i = 1; i < _rows; i++)
                {
                    usersEntities.Workers.Add(new Worker()
                    {
                        id_worker = list[i, 0],
                        status = list[i, 1],
                        fio = list[i, 2],
                        login = list[i, 3],
                        pass = list[i, 4],
                        lastenter = list[i, 5],
                        entertype = list[i, 6]
                    });
                }
                try
                {
                    usersEntities.SaveChanges();
                    MessageBox.Show("Успешный импорт");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Worker> allWorkers;
            using (importisrpo2Entities1 usersEntities = new importisrpo2Entities1())
            {
                allWorkers = usersEntities.Workers.ToList().OrderBy(s => s.status).ToList();
            }

            var app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            var workersCategories = allWorkers.GroupBy(s => s.status).ToList();

            foreach (var workers in workersCategories)
            {
                Excel.Worksheet worksheet = app.Worksheets.Add();
                worksheet.Name = Convert.ToString(workers.Key);

                worksheet.Cells[1, 1] = "Код сотрудника";
                worksheet.Cells[1, 2] = "ФИО";
                worksheet.Cells[1, 3] = "Логин";

                int startRowIndex = 2;
                foreach (Worker worker in workers)
                {
                    worksheet.Cells[startRowIndex, 1] = worker.id_worker;
                    worksheet.Cells[startRowIndex, 2] = worker.fio;
                    worksheet.Cells[startRowIndex, 3] = worker.login;
                    startRowIndex++;
                }

                worksheet.Cells[startRowIndex, 1].Formula = $"=COUNT(A2:A{startRowIndex - 1})";
                worksheet.Cells[startRowIndex, 1].Font.Bold = true;

                Excel.Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[startRowIndex, 3]];
                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                worksheet.Columns.AutoFit();
            }

            app.Visible = true;
        }

    }
}
