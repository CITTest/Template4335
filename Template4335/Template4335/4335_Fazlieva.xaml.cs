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
using Word = Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System.IO;

namespace Template4335
{
    /// <summary>
    /// Логика взаимодействия для _4335_Fazlieva.xaml
    /// </summary>
    public partial class _4335_Fazlieva : Window
    {
        public _4335_Fazlieva()
        {
            InitializeComponent();
        }

        private void BnImport_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (3.xlsx)|*.xlsx",
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
            using (ClientsEntities clientsEntities = new ClientsEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    Clients clients = new Clients();
                    clients.Id = i;
                    clients.FullName = list[i, 0];
                    clients.CodeClient = list[i, 1];
                    clients.BirthDate = list[i, 2];
                    clients.Index = list[i, 3];
                    clients.City = list[i, 4];
                    clients.Street = list[i, 5];
                    clients.Home = int.Parse(list[i, 6]);
                    clients.Kvartira = int.Parse(list[i, 7]);
                    clients.E_mail = list[i, 8];
                    clientsEntities.Clients.Add(clients);
                }
                clientsEntities.SaveChanges();
            }
            MessageBox.Show("Успешно");
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Clients> allClients;
            List<string> allStreet;
            using (ClientsEntities clientsEntities = new ClientsEntities())
            {
                allClients = clientsEntities.Clients.ToList().OrderBy(s => s.FullName).ToList();
                allStreet = clientsEntities.Clients.ToList().Select(s => s.Street).Distinct().ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allStreet.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < allStreet.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(allStreet[i]);
                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "ФИО";
                worksheet.Cells[3][startRowIndex] = "E-mail";
                startRowIndex++;
                var clientsCategories = allClients.GroupBy(s => s.Street).ToList();
                foreach (var clients in clientsCategories)
                {
                    if (clients.Key == allStreet[i])
                    {
                        foreach (Clients client in allClients)
                        {
                            if (client.Street == clients.Key)
                            {
                                worksheet.Cells[1][startRowIndex] = client.CodeClient;
                                worksheet.Cells[2][startRowIndex] = client.FullName;
                                worksheet.Cells[3][startRowIndex] = client.E_mail;
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

        private void BnImportJson_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json;*.json",
                Filter = "файл Excel (1.json)|*.json",
                Title = "Выберите файл"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string jsontext = File.ReadAllText(ofd.FileName);
            List<Clients> clients = JsonConvert.DeserializeObject<List<Clients>>(jsontext);
            using (ClientsEntities clientsEntities = new ClientsEntities())
            {
                clientsEntities.Clients.AddRange(clients);
                clientsEntities.SaveChanges();
            }
            MessageBox.Show("Успешно");
        }

        private void BnExportWord_Click(object sender, RoutedEventArgs e)
        {
            List<Clients> allClients;
            List<string> allStreet;

            using (ClientsEntities clientsEntities = new ClientsEntities())
            {
                allClients = clientsEntities.Clients.ToList().OrderBy(s => s.FullName).ToList();
                allStreet = clientsEntities.Clients.ToList().Select(s => s.Street).Distinct().ToList();
                var clientsCategories = allClients.GroupBy(s => s.Street).ToList();
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();
                int i = 0;
                foreach (var group in clientsCategories)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = Convert.ToString(allStreet[i]);
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();
                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table servicesTable = document.Tables.Add(tableRange, group.Count() + 1, 3);
                    servicesTable.Borders.InsideLineStyle = servicesTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    servicesTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    i++;
                    Word.Range cellRange;
                    cellRange = servicesTable.Cell(1, 1).Range;
                    cellRange.Text = "Код клиента";
                    cellRange = servicesTable.Cell(1, 2).Range;
                    cellRange.Text = "ФИО";
                    cellRange = servicesTable.Cell(1, 3).Range;
                    cellRange.Text = "E-mail";
                    servicesTable.Rows[1].Range.Bold = 1;
                    servicesTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    int j = 1;
                    foreach (var currentService in group)
                    {
                        cellRange = servicesTable.Cell(j + 1, 1).Range;
                        cellRange.Text = currentService.CodeClient;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = servicesTable.Cell(j + 1, 2).Range;
                        cellRange.Text = currentService.FullName;
                        cellRange = servicesTable.Cell(j + 1, 3).Range;
                        cellRange.Text = currentService.E_mail;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        j++;
                    }
                    document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }
                app.Visible = true;

            }
        }
    }
}
