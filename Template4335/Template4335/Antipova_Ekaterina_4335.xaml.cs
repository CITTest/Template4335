using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
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
using System.Text.Json;
using System.IO;
using System.Data.Entity.Validation;
using System.Data.SqlClient;
using System.Security.Cryptography;
using System.Data.Entity;
using System.Text.Json.Serialization;
using Newtonsoft.Json;
using Word = Microsoft.Office.Interop.Word;

namespace Template4335
{
    /// <summary>
    /// Логика взаимодействия для Antipova_Ekaterina_4335.xaml
    /// </summary>
    public partial class Antipova_Ekaterina_4335 : Window
    {
        public Antipova_Ekaterina_4335()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e) // import
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx", //Свойство DefaultExt задает расширение имени файла по умолчанию.
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx", // Свойство Filter задает текущую строку фильтра имен файлов, которая определяет варианты, доступные в поле диалогового окна «Сохранитькак файл типа» или «Файлы типа»
                Title = "Выберите файл базы данных" //Свойство Title задает заголовок диалогового окна файла
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list; //создается двумерный массив строкового типа
            Excel.Application ObjWorkExcel = new Excel.Application(); //создается экземпляр класса Excel.Application для начала работы с библиотекой Interop
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName); //создается экземпляр класса Excel.Workbook для загрузки документа формата xlsx с электронными таблицами
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //необходимо выбрать лист xlsx-файла, с которого в дальнейшем будет происходить чтение данных
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell); //определяется последняя ячейка из таблицы, чтобы определить номер последней строки и столбца содержательной части и сохранить их в переменные _columns и _rows
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns]; //для двумерного массива list выделяется _rows строк и _columns столбцов для будущей записи в него данных из xlsx-файла с помощью вложенного цикла
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);// закрывается сессия работы с книгой Excel.Workbook
            ObjWorkExcel.Quit();//реализован выход из Excel
            GC.Collect();

            using (isrpoEntities2 usersEntities = new isrpoEntities2())
            {
               
                for (int i = 1; i < _rows; i++)
                {
                    usersEntities.Users.Add(new Users()
                    {
                    
                        NaimeovanieUslugi = list[i, 1],
                        VidUslugi = list[i, 2],
                        KodUslugi = list[i, 3],
                        Stoimost = Convert.ToInt32(list[i, 4])
                    });
                }
                usersEntities.SaveChanges();
                MessageBox.Show("все успешно");
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e) // export
        {
            using (var context = new isrpoEntities2())
            {
                var services = context.Users.OrderBy(s => s.Stoimost).ToList();
                var groupedServices = services.GroupBy(s => s.VidUslugi);

                var app = new Excel.Application();
                app.SheetsInNewWorkbook = groupedServices.Count();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

                int sheetIndex = 1;
                foreach (var group in groupedServices)
                {
                    Excel.Worksheet worksheet = workbook.Worksheets.Item[sheetIndex];
                    worksheet.Name = group.Key;

                    worksheet.Cells[1, 1] = "Id";
                    worksheet.Cells[1, 2] = "Название услуги";
                    worksheet.Cells[1, 3] = "Вид услуги";
                    worksheet.Cells[1, 4] = "Код услуги";
                    worksheet.Cells[1, 5] = "Стоимость";

                    int rowIndex = 2;
                    foreach (var service in group)
                    {
                        worksheet.Cells[rowIndex, 1] = service.Id;
                        worksheet.Cells[rowIndex, 2] = service.NaimeovanieUslugi;
                        worksheet.Cells[rowIndex, 3] = service.VidUslugi;
                        worksheet.Cells[rowIndex, 4] = service.KodUslugi;
                        worksheet.Cells[rowIndex, 5] = service.Stoimost;

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

        private void Button_Click_2(object sender, RoutedEventArgs e) // импорт json
        {
            OpenFileDialog ofd = new OpenFileDialog() // Открытие диалогового окна для выбора JSON-файла для импорта данных
            {
                DefaultExt = "*.json",
                Filter = "JSON файлы (*.json)|*.json",
                Title = "Выберите JSON файл для импорта данных"
            };
            if (ofd.ShowDialog() == true) // Проверка, был ли выбран файл и нажата кнопка "ОК"
            {
                try
                {
                    string jsonText = File.ReadAllText(ofd.FileName); // Чтение текста из выбранного JSON-файла
                    List<Users> usersData = JsonConvert.DeserializeObject<List<Users>>(jsonText); // Десериализация JSON-текста в список объектов Users

                    using (isrpoEntities2 usersEntities = new isrpoEntities2())   // Создание нового контекста базы данных
                    {
                        usersEntities.Users.AddRange(usersData);  // Добавление данных из JSON в базу данных
                        usersEntities.SaveChanges(); // Сохранение изменений в базе данных
                    }

                    MessageBox.Show("Данные успешно импортированы из JSON файла и сохранены в базе данных.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Произошла ошибка при импорте данных из JSON файла: {ex.Message}");
                }
            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e) // экспорт word
        {
            // Открытие диалогового окна для выбора JSON-файла для экспорта данных
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "JSON файлы (*.json)|*.json",
                Title = "Выберите JSON файл для экспорта данных"
            };

            if (ofd.ShowDialog() == true) // Проверка, был ли выбран файл и нажата кнопка "ОК"
            {
                // Чтение данных из JSON файла
                string json = File.ReadAllText(ofd.FileName);
                List<Users> users = System.Text.Json.JsonSerializer.Deserialize<List<Users>>(json);  // Десериализация JSON-данных в список объектов Users

                // Отсортировать данные по возрастанию стоимости услуг
                var sortedUsers = users.OrderBy(u => u.Stoimost);

                // Создание нового документа Word
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document doc = app.Documents.Add();


                // Добавление отсортированных данных в документ Word
                foreach (var user in sortedUsers)
                {
                    Microsoft.Office.Interop.Word.Paragraph para = doc.Content.Paragraphs.Add();
                    // Добавление информации о каждом пользователе в документ Word
                    para.Range.Text = $"Id: {user.Id}\nНазвание услуги: {user.NaimeovanieUslugi}\nСтоимость: {user.Stoimost}\n";
                }

                // Сохранение документа Word
                doc.SaveAs2(@"D:\isrpo\sorted_output.docx");
                doc.Close();
                app.Quit();
                MessageBox.Show("Данные были экспортированы в документ Word и отсортированы по стоимости услуг");
            }

        }
    }
   
}
