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

namespace PcSborka.Views.Windows
{
    /// <summary>
    /// Логика взаимодействия для PeripherySelectItem.xaml
    /// </summary>
    public partial class PeripherySelectItem : Window
    {
        public PeripherySelectItem()
        {
            InitializeComponent();
        }

        private void ChooseMonitor_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new ChooseCurrentItemShow(8, "cpu");
            nextWindow.ShowDialog();
            ShowDialog();
        }

        private void ChooseMouse_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new ChooseCurrentItemShow(9, "cpu");
            nextWindow.ShowDialog();
            ShowDialog();
        }

        private void ChooseKeyboard_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new ChooseCurrentItemShow(10, "cpu");
            nextWindow.ShowDialog();
            ShowDialog();

        }

        private void Back_button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
