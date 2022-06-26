using PcSborka.Entity;
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

namespace PcSborka.Views.UserControls
{
    /// <summary>
    /// Логика взаимодействия для CurrentOrderControl.xaml
    /// </summary>
    public partial class CurrentOrderControl : UserControl
    {
        public Order OrderID { get; set; }

        public CurrentOrderControl(Order order)
        {
            InitializeComponent();

            OrderID = order;
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            NumberOrder_textBlock.Text = $"Номер заказа: {OrderID.ID}";
            DateOrder_textBlock.Text = $"Дата заказа: {OrderID.DateOrder.ToShortDateString()}";
            if(OrderID.Accepted != false)
            {
                StatusOrder_textBlock.Text = $"Статус заказ: Завершен";
            }
            else
            {
                StatusOrder_textBlock.Text = $"Статус заказ: Новый";
            }
        }
    }
}
