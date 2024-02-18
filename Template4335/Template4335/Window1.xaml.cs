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
                    Order order = new Order
                    {
                        OrderCode = excelRange.Cells[i, 2].Value2.ToString(),
                        CreationDate = DateTime.FromOADate((double)excelRange.Cells[i, 3].Value2),
                        ClientCode = excelRange.Cells[i, 4].Value2.ToString(),
                        Services = excelRange.Cells[i, 5].Value2.ToString()
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
    }
}
