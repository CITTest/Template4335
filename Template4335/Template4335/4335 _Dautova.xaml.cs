using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Serialization;
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
using System.Text.Json;
using System.IO;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Word;

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

        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (5.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            // Объявление переменных и открытие файла Excel
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];// Получаем доступ к первому листу в книге
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);// Находим последнюю заполненную ячейку на листе
            int _columns = (int)lastCell.Column;// Получаем количество столбцов, равное номеру последней заполненной ячейки в строке
            int _rows = (int)lastCell.Row; // Получаем количество строк, равное номеру последней заполненной ячейки в столбце
            list = new string[_rows, _columns];

            for (int j = 0; j < _columns; j++) // Заполнение массива данными из файла Excel
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;

            // Закрытие файла Excel
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);// Закрытие рабочей книги без сохранения изменений           
            ObjWorkExcel.Quit();// Закрытие Excel 
            GC.Collect();// Очистка неуправляемых ресурсов для уменьшения нагрузки на систему

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
                var groupedEmployees = employees.GroupBy(E => E.Role);// Группируем сотрудников по их ролям

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
                    MessageBox.Show("успешно");
                }
            }
        }

        private void BnImport_JSON_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*..json",
                Filter = "файл json (5.json)|*.json",
                Title = "Выберите файл базы данных"
            };

            if (ofd.ShowDialog() == true) // Проверка, был ли выбран файл и нажата кнопка "ОК"
            {
                string jsonText = File.ReadAllText(ofd.FileName);
                List<Employee> employees = JsonConvert.DeserializeObject<List<Employee>>(jsonText);

                using (EmplEntities usersEntities = new EmplEntities())
                {
                    foreach (Employee emp in employees)
                    {

                        usersEntities.Employees.Add(emp);
                    }
                    usersEntities.SaveChanges();
                    MessageBox.Show("Данные успешно сохранены в базу данных");
                }
            }
            else MessageBox.Show("ошибка");
        }

        private void BnExport_JSON_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog()
            {
                DefaultExt = "docx",
                Filter = "Файл Word (*.docx)|*.docx",
                Title = "Сохранить данные в файл"
            };

            if (sfd.ShowDialog() == true) // Проверка, был ли выбран файл и нажата кнопка "ОК"
            {
                using (EmplEntities emplEntities = new EmplEntities())
                {
                    var employees = emplEntities.Employees.OrderBy(E => E.Role).ToList();
                    var groupedEmployees = employees.GroupBy(E => E.Role);

                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    Document doc = wordApp.Documents.Add();

                    foreach (var group in groupedEmployees)
                    {
                        Microsoft.Office.Interop.Word.Paragraph paragraph = doc.Paragraphs.Add();
                        paragraph.Range.Text = group.Key;
                        Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(paragraph.Range, group.Count() + 1, 2);
                        table.Cell(1, 1).Range.Text = "Login";
                        table.Cell(1, 2).Range.Text = "Password";

                        int rowIndex = 2;
                        foreach (var employee in group)
                        {
                            table.Cell(rowIndex, 1).Range.Text = employee.Login;
                            table.Cell(rowIndex, 2).Range.Text = HashPassword(employee.Pass);
                            rowIndex++;
                        }
                        doc.Words.Last.InsertBreak(WdBreakType.wdPageBreak);
                    }

                    doc.SaveAs2(sfd.FileName);
                    doc.Close();
                    wordApp.Quit();
                    MessageBox.Show("успешно");
                }
            }

        }
    }
}
