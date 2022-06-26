using PcSborka.Entity;
using PcSborka.Views.UserControls;
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
    /// Логика взаимодействия для ShowOrdersWindow.xaml
    /// </summary>
    public partial class ShowOrdersWindow : Window
    {
        public static Order CheckOrder { get; set; }
        public CreatePcForThePeopl_dbEntities DBcontext;

        public ShowOrdersWindow()
        {
            InitializeComponent();

            DBcontext = CreatePcForThePeopl_dbEntities.DBContext;
        }

        private void UpdateListView()
        {
            CurrentItem_listView.Items.Clear();

            List<Order> orders = new List<Order>();

            orders = DBcontext.Order.ToList();

            switch (Sorting_comboBox.SelectedIndex)
            {
                case 1:
                    orders = orders.OrderBy(o => o.DateOrder).ToList();
                    break;
                case 2:
                    orders = orders.OrderByDescending(o => o.DateOrder).ToList();
                    break;
                case 3:
                    orders = orders.Where(l => l.Accepted == true).ToList();
                    break;
                case 4:
                    orders = orders.Where(l => l.Accepted == false).ToList();
                    break;
            }

            foreach (Order order in orders)
            {
                CurrentItem_listView.Items.Add(new CurrentOrderControl(order));
            }
        }

        private void Sorting_comboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateListView();
        }

        private void CurrentItem_listView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void CurrentItem_listView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Order order = ((CurrentOrderControl)CurrentItem_listView.SelectedItem).OrderID;
            new FinishToOrderWindow(order).ShowDialog();
        }

        private void BackWindow_button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Sorting_comboBox.Items.Add("Все");
            Sorting_comboBox.Items.Add("По дате ↑");
            Sorting_comboBox.Items.Add("По дате ↓");
            Sorting_comboBox.Items.Add("По статусу:Завершен");
            Sorting_comboBox.Items.Add("По статусу:Новый");

            UpdateListView();
        }

        private void DeleteOrder_button_Click(object sender, RoutedEventArgs e)
        {
            if(CurrentItem_listView.SelectedItem !=null)
            {
                Order order = ((CurrentOrderControl)CurrentItem_listView.SelectedItem).OrderID;
                DBcontext.Order.Remove(order);
                DBcontext.SaveChanges();
                UpdateListView();
            }
            else
            {
                MessageBox.Show("Выберите заказ пожалуйста","Предупреждение", MessageBoxButton.OK,MessageBoxImage.Hand);
            }
        }
    }
}
