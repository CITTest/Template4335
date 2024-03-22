using System;
using Microsoft.Win32;
using System.Windows;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Data.SqlClient;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Text.Json;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    public partial class MainWindow : System.Windows.Window
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
                        excelEntities.Users.Add(new Users()
                        {
                            Client_ID = Convert.ToInt32(list[i, 0]),
                            Name = list[i, 1],
                            Date_Birth = list[i, 2],
                            Index_ = list[i, 3],
                            City = list[i, 4],
                            Street = list[i, 5],
                            House = Convert.ToInt32(list[i, 6]),
                            Flat = Convert.ToInt32(list[i, 7]),
                            Email = list[i, 8]
                        });
                    }
                    excelEntities.SaveChanges();
                }
            }
        }

        public int DateTimeConvert(string Date_Birth)
        {
            var dateOfBirth = DateTime.ParseExact(Date_Birth, "dd.MM.yyyy", CultureInfo.InvariantCulture);
            var currentAge = DateTime.Now.Year - dateOfBirth.Year;
            if (DateTime.Now.DayOfYear < dateOfBirth.DayOfYear)
                currentAge++;
            return currentAge;
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Users> allEntities;
            List<string> allType = new List<string>() { "Категория 1", "Категория 2", "Категория 3" };
            using (ExcelLabEntities ExcelEntities = new ExcelLabEntities())
            {
                allEntities = ExcelEntities.Users.ToList().OrderBy(s => s.Client_ID).ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allType.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < allType.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(allType[i]);
                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "ФИО";
                worksheet.Cells[3][startRowIndex] = "Email";
                startRowIndex++;
                foreach (Users service in allEntities)
                {
                    if (DateTimeConvert(service.Date_Birth) >= 20 && DateTimeConvert(service.Date_Birth) <= 29 && allType[i] == "Категория 1")
                    {
                        worksheet.Cells[1][startRowIndex] = service.Client_ID;
                        worksheet.Cells[2][startRowIndex] = service.Name;
                        worksheet.Cells[3][startRowIndex] = service.Email;
                        startRowIndex++;
                    }
                    else if (DateTimeConvert(service.Date_Birth) >= 30 && DateTimeConvert(service.Date_Birth) <= 39 && allType[i] == "Категория 2")
                    {
                        worksheet.Cells[1][startRowIndex] = service.Client_ID;
                        worksheet.Cells[2][startRowIndex] = service.Name;
                        worksheet.Cells[3][startRowIndex] = service.Email;
                        startRowIndex++;
                    }
                    else if (DateTimeConvert(service.Date_Birth) >= 40 && allType[i] == "Категория 3")
                    {
                        worksheet.Cells[1][startRowIndex] = service.Client_ID;
                        worksheet.Cells[2][startRowIndex] = service.Name;
                        worksheet.Cells[3][startRowIndex] = service.Email;
                        startRowIndex++;
                    }
                }
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }

        private void Btn2Import_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json;*.json",
                Filter = "файл json (3.json)|*.json",
                Title = "Выберите файл json"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            string jsonData = File.ReadAllText(ofd.FileName);

            List<Users1> data = JsonSerializer.Deserialize<List<Users1>>(jsonData);

            using (ExcelLabEntities JsonEntities = new ExcelLabEntities())
            {
                foreach (var item in data)
                {
                    JsonEntities.Users1.Add(new Users1()
                    {
                        Id = item.Id,
                        CodeClient = item.CodeClient,
                        FullName = item.FullName,
                        BirthDate = item.BirthDate,
                        Index = item.Index,
                        City = item.City,
                        Street = item.Street,
                        Home = Convert.ToInt32(item.Home),
                        Kvartira = Convert.ToInt32(item.Kvartira),
                        E_mail = item.E_mail
                    });
                }
                JsonEntities.SaveChanges();
            }
        }

        private void Btn2Export_Click(object sender, RoutedEventArgs e)
        {
            List<Users1> allUsers;
            List<string> allType = new List<string>() { "Категория 1", "Категория 2", "Категория 3" };
            List<Users1> sortedUsers = new List<Users1>();
            using (ExcelLabEntities ExcelEntities = new ExcelLabEntities())
            {
                allUsers = ExcelEntities.Users1.ToList();
            }

            var app = new Word.Application();
            Word.Document document = app.Documents.Add();

            foreach (var type in allType)
            {
                var paragraph = document.Paragraphs.Add();
                paragraph.Range.Text = type;
                paragraph.set_Style("Заголовок 1");
                paragraph.Range.InsertParagraphAfter();

                foreach (Users1 user in allUsers)
                {
                    if (DateTimeConvert(user.BirthDate) >= 20 && DateTimeConvert(user.BirthDate) <= 29 && type == "Категория 1")
                    {
                        sortedUsers.Add(user);
                    }
                    else if (DateTimeConvert(user.BirthDate) >= 30 && DateTimeConvert(user.BirthDate) <= 39 && type == "Категория 2")
                    {
                        sortedUsers.Add(user);
                    }
                    else if (DateTimeConvert(user.BirthDate) >= 40 && type == "Категория 3")
                    {
                        sortedUsers.Add(user);
                    }
                }

                var tableParagraph = document.Paragraphs.Add();
                var tableRange = tableParagraph.Range;
                var usersTable = document.Tables.Add(tableRange, sortedUsers.Count + 1, 3);
                usersTable.Borders.InsideLineStyle = usersTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                usersTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                foreach (Users1 user in sortedUsers)
                {
                    string[] columnHeaders = { "Код клиента", "ФИО", "Email" };
                    for (int i = 0; i < columnHeaders.Length; i++)
                    {
                        usersTable.Cell(1, i + 1).Range.Text = columnHeaders[i];
                        usersTable.Rows[1].Range.Bold = 1;
                        usersTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }

                    for (int i = 0; i < sortedUsers.Count; i++)
                    {
                        var order = sortedUsers[i];
                        usersTable.Cell(i + 2, 1).Range.Text = order.CodeClient.ToString();
                        usersTable.Cell(i + 2, 2).Range.Text = order.FullName;
                        usersTable.Cell(i + 2, 3).Range.Text = order.E_mail;
                    }
                }
                document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                sortedUsers.Clear();
            }
            app.Visible = true;
        }
    }
}
