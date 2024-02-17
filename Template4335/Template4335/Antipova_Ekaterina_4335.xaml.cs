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

        }
    }
}
