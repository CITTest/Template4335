using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


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
                    services.IdServices = int.Parse(list[i, 0]);
                    services.NameServices = list[i, 1];
                    services.TypeOfService = list[i, 2];
                    services.CodeService = list[i, 3];
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
                allType = lrEntities.Services.ToList().Select(s => s.TypeOfService).Distinct().ToList();
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
                var servicesCategories = allServices.GroupBy(s => s.TypeOfService).ToList();
                foreach (var services in servicesCategories)
                {
                    if (services.Key == allType[i])
                    {
                        foreach (Services service in allServices)
                        {
                            if (service.TypeOfService == services.Key)
                            {
                                worksheet.Cells[1][startRowIndex] = service.IdServices;
                                worksheet.Cells[2][startRowIndex] = service.NameServices;
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
            string text = File.ReadAllText(ofd.FileName);
            List<Services> services = JsonConvert.DeserializeObject<List<Services>>(text);

            using (LR3Entities lrEntities = new LR3Entities())
            {
                lrEntities.Services.AddRange(services);
                lrEntities.SaveChanges();
            }
            MessageBox.Show("Импорт данных прошел успешно");
        }

        private void BnExportWord_Click(object sender, RoutedEventArgs e)
        {
            List<Services> allServices;
            List<string> allType;
            using (LR3Entities lrEntities = new LR3Entities())
            {
                allServices = lrEntities.Services.ToList().OrderBy(s => s.Cost).ToList();
                allType = lrEntities.Services.ToList().Select(s => s.TypeOfService).Distinct().ToList();
                var servicesCategories = allServices.GroupBy(s =>s.TypeOfService).ToList();
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();
                int i = 0;
                foreach (var group in servicesCategories)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = Convert.ToString(allType[i]);
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();
                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table servicesTable =  document.Tables.Add(tableRange, group.Count() + 1, 3);
                    servicesTable.Borders.InsideLineStyle = servicesTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    servicesTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    i++;
                    Word.Range cellRange;
                    cellRange = servicesTable.Cell(1, 1).Range;
                    cellRange.Text = "ID";
                    cellRange = servicesTable.Cell(1, 2).Range;
                    cellRange.Text = "Название услуги";
                    cellRange = servicesTable.Cell(1, 3).Range;
                    cellRange.Text = "Стоимость";
                    servicesTable.Rows[1].Range.Bold = 1;
                    servicesTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    int j = 1;
                    foreach (var currentService in group)
                    {
                        cellRange = servicesTable.Cell(j + 1, 1).Range;
                        cellRange.Text = currentService.IdServices.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = servicesTable.Cell(j + 1, 2).Range;
                        cellRange.Text = currentService.NameServices;
                        cellRange = servicesTable.Cell(j + 1, 3).Range;
                        cellRange.Text = currentService.Cost.ToString();
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
