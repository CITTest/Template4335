using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Security.Cryptography;
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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4335
{
    /// <summary>
    /// Логика взаимодействия для _4335__Dautova.xaml
    /// </summary>
    public partial class _4335__Dautova : System.Windows.Window
    {
        public _4335__Dautova()
        {
            InitializeComponent();
        }

        private void BnImport_Click(object sender,
        RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (5.xlsx)|*.xlsx",
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

            using (EmplEntities usersEntities = new EmplEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    usersEntities.Employees.Add(new Employee()
                    {
                        Role = list[i, 0],
                        FullName = list[i, 1],
                        Login = list[i, 2],
                        Pass = list[i, 3]
                    });
                }
                usersEntities.SaveChanges();
                MessageBox.Show("успешно");
            }
        }

        private string HashPassword(string password)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(password);
                byte[] hash = sha256.ComputeHash(bytes);
                return Convert.ToBase64String(hash);
            }
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            using (EmplEntities emplEntities = new EmplEntities())
            {
                var employees = emplEntities.Employees.OrderBy(E => E.Role).ToList(); //списка всех сотрудников, отсортированных по роли
                var groupedEmployees = employees.GroupBy(E => E.Role);

                var app = new Excel.Application();
                app.SheetsInNewWorkbook = groupedEmployees.Count();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
                int sheetIndex = 1;

                foreach (var group in groupedEmployees)
                {
                    Excel.Worksheet worksheet = workbook.Worksheets.Item[sheetIndex];
                    worksheet.Name = group.Key;
                    worksheet.Cells[1, 1] = "Login";
                    worksheet.Cells[1, 2] = "Password";

                    int rowIndex = 2;
                    foreach (var employee in group)
                    {
                        worksheet.Cells[rowIndex, 1] = employee.Login;
                        worksheet.Cells[rowIndex, 2] = HashPassword(employee.Pass);
                        rowIndex++;
                    }
                    sheetIndex++;
                }

                SaveFileDialog sfd = new SaveFileDialog()
                {
                    DefaultExt = "*.xlsx",
                    Filter = "Файл Excel (*.xlsx)|*.xlsx",
                    Title = "Сохранить данные в файл"
                };

                if (sfd.ShowDialog() == true)
                {
                    workbook.SaveAs(sfd.FileName);
                    workbook.Close();
                    app.Quit();
                    MessageBox.Show("Данные успешно экспортированы.");
                }
            }
        }
    }
}
