using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4335
{
    public partial class Window1 : Window
    {
        public Window1()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Указываем путь к файлу Excel
            string filePath = "C:/Users/sljus/OneDrive/Рабочий стол/2.xlsx";

            // Создаем экземпляр DataLoader
            DataLoader dataLoader = new DataLoader();

            // Загружаем данные из Excel и сохраняем в базу данных
            dataLoader.LoadDataFromExcel(filePath);

            // Выводим сообщение об успешном импорте данных
            MessageBox.Show("Данные успешно импортированы из файла Excel и сохранены в базу данных.");
        }

        public class OrderContext : DbContext
        {
            public DbSet<Order> Orders { get; set; }
        }

        public class Order
        {
            public int Id { get; set; }
            public string OrderCode { get; set; }
            public DateTime CreationDate { get; set; }
            public string ClientCode { get; set; }
            public string Services { get; set; }
            public string Status { get; set; }
            public DateTime ClosingDate { get; set; }
            public string RentTime { get; set; }
        }

        public class DataLoader
        {
            public void LoadDataFromExcel(string filePath)
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath);
                Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;

                int rowCount = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;

                List<Order> orders = new List<Order>();

                for (int i = 2; i <= rowCount; i++)
                {
                    string orderCode = excelRange.Cells[i, 2]?.Value2?.ToString();
                    DateTime creationDate = DateTime.FromOADate((double)(excelRange.Cells[i, 3]?.Value2 ?? 0));
                    string clientCode = excelRange.Cells[i, 5]?.Value2?.ToString();
                    string services = excelRange.Cells[i, 6]?.Value2?.ToString();
                    string status = excelRange.Cells[i, 7]?.Value2?.ToString();
                    DateTime closingDate = DateTime.FromOADate((double)(excelRange.Cells[i, 8]?.Value2 ?? 0));
                    string rentTime = excelRange.Cells[i, 9]?.Value2?.ToString();

                    Order order = new Order
                    {
                        OrderCode = orderCode,
                        CreationDate = creationDate,
                        ClientCode = clientCode,
                        Services = services,
                        Status = status,
                        ClosingDate = closingDate,
                        RentTime = rentTime
                    };
                    orders.Add(order);
                }

                excelWorkbook.Close();
                excelApp.Quit();

                using (var context = new OrderContext())
                {
                    context.Orders.AddRange(orders);
                    context.SaveChanges();
                }
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            using (var context = new OrderContext())
            {
                // Получаем данные из базы данных, отсортированные по времени проката
                var ordersByRentTime = context.Orders.OrderBy(o => o.RentTime).ToList();

                // Создаем новый файл Excel
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();

                // Получаем уникальные значения времени проката для разделения на категории
                var uniqueRentTimes = ordersByRentTime.Select(o => o.RentTime).Distinct();

                foreach (var rentTime in uniqueRentTimes)
                {
                    // Создаем новый лист Excel для текущей категории
                    Excel._Worksheet excelWorksheet = excelWorkbook.Sheets.Add();
                    excelWorksheet.Name = $"RentTime_{rentTime}";

                    // Фильтруем данные для текущей категории
                    var ordersInCategory = ordersByRentTime.Where(o => o.RentTime == rentTime).ToList();

                    // Записываем данные в лист Excel
                    for (int i = 0; i < ordersInCategory.Count; i++)
                    {
                        excelWorksheet.Cells[i + 1, 1] = ordersInCategory[i].Id;
                        excelWorksheet.Cells[i + 1, 2] = ordersInCategory[i].OrderCode;


                        excelWorksheet.Cells[i + 1, 3] = ordersInCategory[i].CreationDate;

                        excelWorksheet.Cells[i + 1, 4] = ordersInCategory[i].ClientCode;
                        excelWorksheet.Cells[i + 1, 5] = ordersInCategory[i].Services;
                    }
                }

                // Сохраняем новый файл Excel
                string exportFilePath = "C:\\Users\\sljus\\source\\repos\\ExportedData.xlsx";
                excelWorkbook.SaveAs(exportFilePath);
                excelWorkbook.Close();
                excelApp.Quit();

                // Выводим сообщение об успешном экспорте данных
                MessageBox.Show("Данные успешно экспортированы в новый файл Excel.");
            }

        }

    }
}

