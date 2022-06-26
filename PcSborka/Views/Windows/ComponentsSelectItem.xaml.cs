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
    /// Логика взаимодействия для ComponentsSelectItem.xaml
    /// </summary>
    public partial class ComponentsSelectItem : Window
    {
        public ComponentsSelectItem()
        {
            InitializeComponent();
        }

        private void Back_button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ChooseCPU_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new ChooseCurrentItemShow(1, "cpu");
            nextWindow.ShowDialog();
            ShowDialog();
        }

        private void ChooseMotherBoard_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new ChooseCurrentItemShow(2, "cpu");
            nextWindow.ShowDialog();
            ShowDialog();
        }

        private void ChooseCase_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new ChooseCurrentItemShow(3, "cpu");
            nextWindow.ShowDialog();
            ShowDialog();
        }

        private void ChooseGPU_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new ChooseCurrentItemShow(4, "cpu");
            nextWindow.ShowDialog();
            ShowDialog();
        }

        private void ChooseCooler_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new ChooseCurrentItemShow(5, "cpu");
            nextWindow.ShowDialog();
            ShowDialog();
        }

        private void ChooseRAM_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new ChooseCurrentItemShow(6, "cpu");
            nextWindow.ShowDialog();
            ShowDialog();
        }

        private void ChoosePowerSupply_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new ChooseCurrentItemShow(7, "cpu");
            nextWindow.ShowDialog();
            ShowDialog();
        }
    }
}
