using Microsoft.Office.Interop.Excel;
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

namespace Template4335
{
    /// <summary>
    /// Логика взаимодействия для _4335__Dautova.xaml
    /// </summary>
    public partial class _4335__Dautova : Window
    {
        public _4335__Dautova()
        {
            InitializeComponent();
        }

        private void BnImport_Click(object sender,
        RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (5.xlsx)|*.xlsx",
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

            using (EmplEntities usersEntities = new EmplEntities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    usersEntities.Employees.Add(new Employee()
                    {
                        Role = list[i, 0],
                        FullName = list[i,1],
                        Login = list[i, 2],
                        Pass = list[i, 3]
                    });
                }
                usersEntities.SaveChanges();
            }
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Student> allStudents;
            List<Group> allGroups;
            using (UsersEntities usersEntities = new UsersEntities())
            {
                allStudents =
                usersEntities.Students.ToList().OrderBy(s =>
                s.Name).ToList();
                allGroups = usersEntities.Groups.ToList().OrderBy(g =>
                g.NumberGroup).ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allGroups.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < allGroups.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i +
                1];
                worksheet.Name =
                Convert.ToString(allGroups[i].NumberGroup);
            }
            for (int i = 0; i < allGroups.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(allGroups[i].NumberGroup);
                int startRowIndex = 1;
                for (int i = 0; i < allGroups.Count(); i++)
                {
                    Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                    worksheet.Name = Convert.ToString(allGroups[i].NumberGroup);
                    worksheet.Cells[1][startRowIndex] = "Порядковый номер";
                worksheet.Cells[2][startRowIndex] = "ФИО студента";
                    startRowIndex++;
                }

            }
            var studentsCategories = allStudents.GroupBy(s => s.Group.NumberGroup).ToList();
            foreach (var students in studentsCategories)
            {
                if (students.Key == allGroups[i].Id)
                {
                    Excel.Range headerRange =
                    worksheet.Range[worksheet.Cells[1][1],
                    worksheet.Cells[2][1]];
                    headerRange.Merge();
                    headerRange.Value = allGroups[i].NumberGroup;
                    headerRange.HorizontalAlignment =
                    Excel.XlHAlign.xlHAlignCenter;
                    headerRange.Font.Italic = true;
                    startRowIndex++;
                }
                else
                {
                    continue;
                }
            }
            foreach (var students in studentsCategories)
            {
                if (students.Key == allGroups[i].Id)
                {
                    Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1],
                    worksheet.Cells[2][1]];
                    headerRange.Merge();
                    headerRange.Value = allGroups[i].NumberGroup;
                    headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    headerRange.Font.Italic = true;
                    startRowIndex++;
                    foreach (Student student in allStudents)
                    {
                        if (student.NumberGroupId == students.Key)
                        {
                            worksheet.Cells[1][startRowIndex] =
                            student.Id;
                            worksheet.Cells[2][startRowIndex] =
                            student.Name;
                            startRowIndex++;
                        }
                    }
                    worksheet.Cells[1][startRowIndex].Formula =
                    $"=СЧЁТ(A3:A{startRowIndex - 1})";
                    worksheet.Cells[1][startRowIndex].Font.Bold =
                    true;
                }
                else
                {
                    continue;
                }
            }
        }
    }
}
