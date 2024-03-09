
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
using Newtonsoft.Json;
using System.IO;
using Word = Microsoft.Office.Interop.Word;




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

        private void Bnlb5(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "JSON файлы (*.json)|*.json",
                Title = "Выберите файл JSON для импорта данных"
            };

            if (ofd.ShowDialog() == true)
            {
                try
                {
                    string jsonContent = File.ReadAllText(ofd.FileName);
                    List<MyDataModel> data = JsonConvert.DeserializeObject<List<MyDataModel>>(jsonContent);

                    using (vtorEntities1 vtorEntities = new vtorEntities1())
                    {
                        foreach (var item in data)
                        {
                            Uslugi uslugi = new Uslugi();
                            uslugi.Id = item.Id;
                            uslugi.CodeZakaz = item.CodeOrder;
                            uslugi.Date = item.CreateDate;
                            uslugi.TimeZakaz = item.CreateTime;
                            uslugi.CodeClient = item.CodeClient;
                            uslugi.Uslugi1 = item.Services;
                            uslugi.Status = item.Status;
                            uslugi.DateOff = item.ClosedDate;
                            uslugi.TimeProkat = item.ProkatTime;
                            vtorEntities.Uslugi.Add(uslugi);
                            // Продолжайте заполнение полей объекта Uslugi в соответствии с вашей моделью данных

                            vtorEntities.Uslugi.Add(uslugi);
                        }
                        vtorEntities.SaveChanges();
                        MessageBox.Show("Импорт данных из JSON файла успешно завершен.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при импорте данных: {ex.Message}");
                }
            }
        }

        // Модель данных для десериализации JSON
        public class MyDataModel
        {
            public int Id { get; set; }
            public string CodeOrder { get; set; }
            public string CreateDate { get; set; }
            public string CreateTime { get; set; }
            public string CodeClient { get; set; }
            public string Services { get; set; }
            public string Status { get; set; }
            public string ClosedDate { get; set; }
            public string ProkatTime { get; set; }
        }



        private void Bnls5js(object sender, RoutedEventArgs e)
        {
            List<Uslugi> allusl;
            List<string> allDate;
            using (vtorEntities1 vtorEntities = new vtorEntities1())
            {
                allusl = vtorEntities.Uslugi.ToList();
                allDate = vtorEntities.Uslugi.ToList().Select(s => s.Date).Distinct().ToList();
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();
                var servicesCategories = allusl.GroupBy(s => s.Date).ToList();
                foreach (var group in servicesCategories)
                {
                    Word.Paragraph paragraph =
                    document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = Convert.ToString(allusl.Where(g =>
g.Date == group.Key).FirstOrDefault().Date);
paragraph.set_Style("Заголовок 1");
range.InsertParagraphAfter();
                    Word.Paragraph tableParagraph =
document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table studentsTable =
                    document.Tables.Add(tableRange, group.Count() + 1, 4);
                    studentsTable.Borders.InsideLineStyle =
                    studentsTable.Borders.OutsideLineStyle =
                    Word.WdLineStyle.wdLineStyleSingle;
                    studentsTable.Range.Cells.VerticalAlignment =
                    Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    Word.Range cellRange;
                    cellRange = studentsTable.Cell(1, 1).Range;
                    cellRange.Text = "ID";
                    cellRange = studentsTable.Cell(1, 2).Range;
                    cellRange.Text = "Код заказа";
                    cellRange = studentsTable.Cell(1, 3).Range;
                    cellRange.Text = "Код клиента";
                    cellRange = studentsTable.Cell(1, 4).Range;
                    cellRange.Text = "Услуги";
                    studentsTable.Rows[1].Range.Bold = 1;
                    studentsTable.Rows[1].Range.ParagraphFormat.Alignment =
                    Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    
                    int i = 1;
                    foreach (var currentStudent in group)
                    {
                        cellRange = studentsTable.Cell(i + 1, 1).Range;
                        cellRange.Text = currentStudent.Id.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        cellRange = studentsTable.Cell(i + 1, 2).Range;
                        cellRange.Text = currentStudent.CodeZakaz;
                        cellRange.ParagraphFormat.Alignment =Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        cellRange = studentsTable.Cell(i + 1, 3).Range;
                        cellRange.Text = currentStudent.CodeClient;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        cellRange = studentsTable.Cell(i + 1, 4).Range;
                        cellRange.Text = currentStudent.Uslugi1;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        i++;

                        document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                    }
                    app.Visible = true;
                    document.SaveAs2(@"D:\outputFileWord.docx");
                    document.SaveAs2(@"D:\outputFilePdf.pdf",
                    Word.WdExportFormat.wdExportFormatPDF);
                }


                }


            }
        }
    }








