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
using PcSborka.Entity;

namespace PcSborka.Views.Windows
{
    /// <summary>
    /// Логика взаимодействия для MenuForTheEmployeer.xaml
    /// </summary>
    public partial class MenuForTheEmployeer : Window
    {
        public MenuForTheEmployeer(Employeer user)
        {
            InitializeComponent();
        }

        private void Back_button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void CheckCompo_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new ComponentsSelectItem();
            nextWindow.ShowDialog();
            Show();
        }

        private void CheckPeriphery_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new PeripherySelectItem();
            nextWindow.ShowDialog();
            Show();
        }

        private void CheckReadyComputers_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new ChooseCurrentItemShow(11,"anime");
            nextWindow.ShowDialog();
            Show();
        }

        private void CheckReadyPeriphery_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new ChooseCurrentItemShow(12, "anime");
            nextWindow.ShowDialog();
            Show();
        }

        private void CheckOrder_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new ShowOrdersWindow();
            nextWindow.ShowDialog();
            Show();
        }

        private void CheckEmployeers_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new EmployeersShowWindow();
            nextWindow.ShowDialog();
            Show();
        }
    }
}
