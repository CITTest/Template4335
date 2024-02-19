﻿using Microsoft.Win32;
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
    }
}
