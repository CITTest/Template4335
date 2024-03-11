using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Text.Json;
using System.IO;
using System.Xml.Linq;
using System.Xaml;


namespace Template4335
{
    /// <summary>
    /// Логика взаимодействия для _4335_BARYSHEVA.xaml
    /// </summary>
    public partial class _4335_BARYSHEVA : System.Windows.Window
    {
        public _4335_BARYSHEVA()
        {
            InitializeComponent();
        }

        private void exportButton_Click(object sender, RoutedEventArgs e)
        {
            List<ORDERS> allOrders;
            List<string> allTimeProkats;
            using (orderszakaz2Entities orderzakazEntities = new orderszakaz2Entities())
            {
                allOrders = orderzakazEntities.ORDERS.ToList();
                allTimeProkats = orderzakazEntities.ORDERS.Select(o => o.ZakazProkatTime).Distinct().ToList();
            }

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allTimeProkats.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < allTimeProkats.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(allTimeProkats[i]);
                //worksheet.Name = Convert.ToString(allTimeProkats[i]) + "_" + i;

                worksheet.Cells[1][startRowIndex] = "Id";
                worksheet.Cells[2][startRowIndex] = "Код заказа";
                worksheet.Cells[3][startRowIndex] = "Дата создания";
                worksheet.Cells[4][startRowIndex] = "Код клиента";
                worksheet.Cells[5][startRowIndex] = "Услуги";
                startRowIndex++;

                var ordersCategories = allOrders.GroupBy(s => s.ZakazProkatTime).ToList();
                foreach (var orders in ordersCategories)
                {
                    if (orders.Key == allTimeProkats[i])
                    {
                        foreach (ORDERS order in allOrders)
                        {
                            if (order.ZakazProkatTime == orders.Key)
                            {
                                worksheet.Cells[1][startRowIndex] = order.ZakazId;
                                worksheet.Cells[2][startRowIndex] = order.ZakazCode;
                                worksheet.Cells[3][startRowIndex] = order.ZakazCreationDate;
                                worksheet.Cells[4][startRowIndex] = order.ClientCode;
                                worksheet.Cells[5][startRowIndex] = order.ZakazUslugi;
                                startRowIndex++;
                            }

                        }
                    }
                    else
                    {
                        continue;
                    }

                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][startRowIndex - 1]];
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();

            }
            app.Visible = true;

        }

        private void importButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
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

            for (int i = 0; i < _rows; i++)
            {
                for (int j = 0; j < _columns; j++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (orderszakaz2Entities orderzakazEntities = new orderszakaz2Entities())
            {


                for (int l = 1; l < 51; l++)
                {
                    ORDERS orders = new ORDERS();
                    orders.ZakazId = Convert.ToInt32(list[l, 0]);
                    orders.ZakazCode = list[l, 1];
                    orders.ZakazCreationDate = list[l, 2];
                    orders.ZakazTime = list[l, 3];
                    orders.ClientCode = list[l, 4];
                    orders.ZakazUslugi = list[l, 5];
                    orders.ZakazStatus = list[l, 6];
                    orders.ZakazClosureDate = list[l, 7];
                    orders.ZakazProkatTime = list[l, 8];

                    //orders.ZakazProkatTime = Regex.Replace(list[l, 8], @"\D", "");

                    orderzakazEntities.ORDERS.Add(orders);
                }
                orderzakazEntities.SaveChanges();
                MessageBox.Show("Импорт данных прошел успешно");
            }

        }

        private void importJSONButton_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "JSON файлы (*.json)|*.json",
                Title = "Выберите файл JSON"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            string jsonContent = File.ReadAllText(ofd.FileName);
            List<ORDERS> dataList = JsonSerializer.Deserialize<List<ORDERS>>(jsonContent);

            using (orderszakaz2Entities orderzakazEntities = new orderszakaz2Entities())
            {
                foreach (var data in dataList)
                {
                    ORDERS orders = new ORDERS();

                    orders.ZakazId = data.ZakazId;
                    orders.ZakazCode = data.ZakazCode;
                    orders.ZakazCreationDate = data.ZakazCreationDate;
                    orders.ZakazTime = data.ZakazTime;
                    orders.ClientCode = data.ClientCode;
                    orders.ZakazUslugi = data.ZakazUslugi;
                    orders.ZakazStatus = data.ZakazStatus;
                    orders.ZakazClosureDate = data.ZakazClosureDate;
                    orders.ZakazProkatTime = data.ZakazProkatTime;

                    orderzakazEntities.ORDERS.Add(orders);
                }

                orderzakazEntities.SaveChanges();
                MessageBox.Show("Импорт данных из JSON прошел успешно");
            }

        }

        private void exportWordButton_Click(object sender, RoutedEventArgs e)
        {
            List<ORDERS> allOrders;
            List<string> allTimeProkats;
            using (orderszakaz2Entities orderzakazEntities = new orderszakaz2Entities())
            {
                allOrders = orderzakazEntities.ORDERS.ToList();
                allTimeProkats = orderzakazEntities.ORDERS.Select(o => o.ZakazProkatTime).Distinct().ToList();
                var ordersCategories = allOrders.GroupBy(s => s.ZakazProkatTime).ToList();

                var app = new Word.Application();
                Word.Document document = app.Documents.Add();

                DateTime firstOrderDate = allOrders.Min(o => DateTime.Parse(o.ZakazCreationDate));
                DateTime lastOrderDate = allOrders.Max(o => DateTime.Parse(o.ZakazCreationDate));

                int i =  0;
                foreach (var timeProkat in ordersCategories)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;

                    range.Text = Convert.ToString(allTimeProkats.Where(t => t == timeProkat.Key).FirstOrDefault());

                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table ordersTable = document.Tables.Add(tableRange, timeProkat.Count() + 1, 5);
                    ordersTable.Borders.InsideLineStyle = ordersTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    ordersTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    i++;

                    Word.Range cellRange;
                    cellRange = ordersTable.Cell(1, 1).Range;
                    cellRange.Text = "Id";
                    cellRange = ordersTable.Cell(1, 2).Range;
                    cellRange.Text = "Код заказа";
                    cellRange = ordersTable.Cell(1, 3).Range;
                    cellRange.Text = "Дата создания";
                    cellRange = ordersTable.Cell(1, 4).Range;
                    cellRange.Text = "Код клиента";
                    cellRange = ordersTable.Cell(1, 5).Range;
                    cellRange.Text = "Услуги";
                    ordersTable.Rows[1].Range.Bold = 1;
                    ordersTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    int j = 1;
                    foreach (var currentOrder in timeProkat)
                    {
                        cellRange = ordersTable.Cell(j + 1, 1).Range;
                        cellRange.Text = currentOrder.ZakazId.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = ordersTable.Cell(j + 1, 2).Range;
                        cellRange.Text = currentOrder.ZakazCode;
                        cellRange = ordersTable.Cell(j + 1, 3).Range;
                        cellRange.Text = currentOrder.ZakazCreationDate;
                        cellRange = ordersTable.Cell(j + 1, 4).Range;
                        cellRange.Text = currentOrder.ClientCode;
                        cellRange = ordersTable.Cell(j + 1, 5).Range;
                        cellRange.Text = currentOrder.ZakazUslugi;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        j++;
                    }

                    document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }

                Word.Paragraph overallDatesParagraph = document.Paragraphs.Add();
                overallDatesParagraph.Range.Text = $"Дата первого заказа: {firstOrderDate.ToShortDateString()}, Дата последнего заказа: {lastOrderDate.ToShortDateString()}";
                overallDatesParagraph.Format.SpaceAfter = 12;
                overallDatesParagraph.Range.InsertParagraphAfter();

                app.Visible = true;

            }
        }
    }
}
