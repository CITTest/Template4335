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
    }
}
