using PcSborka.Entity;
using PcSborka.Views.Windows;
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

namespace PcSborka
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public CreatePcForThePeopl_dbEntities context;
        public Employeer user;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Window nextWindow = new Authorization();
            nextWindow.ShowDialog();
            Close();
        }

        private void GoToCreateOrderWindow_button_Click(object sender, RoutedEventArgs e)
        {
            Window nextWindow = new OrderCreate();
            nextWindow.ShowDialog();
        }

        private void CheckOrder_button_Click(object sender, RoutedEventArgs e)
        {
            Window nextWindow = new InputEmailWindow();
            nextWindow.ShowDialog();
            Close();
        }
    }
}
