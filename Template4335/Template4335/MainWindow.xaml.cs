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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Template4335
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Porfirev_4335 window;
        public MainWindow()
        {
            InitializeComponent();
        }
        

        private void Porfirev_4335(object sender, RoutedEventArgs e)
        {
            this.window = new Porfirev_4335();
            window.Show();
            this.Close();
        }
    }
}
