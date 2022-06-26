using PcSborka.Entity;
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

namespace PcSborka.Views.Windows
{
    /// <summary>
    /// Логика взаимодействия для FinalStageOfOrder.xaml
    /// </summary>
    public partial class FinalStageOfOrder : Window
    {
        public static Order OrderReadyNot { get; set; } = new Order();
        public CreatePcForThePeopl_dbEntities Dbcontext;
        public static bool isChecked = false;

        public FinalStageOfOrder()
        {
            InitializeComponent();

            Dbcontext = CreatePcForThePeopl_dbEntities.DBContext;
        }

        private void Phone_textBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !(Char.IsDigit(e.Text, 0));
        }

        private void Phone_textBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                e.Handled = true;
            }
        }

        private void CreateOrder_button_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(Email_textBox.Text) && !string.IsNullOrWhiteSpace(Phone_textBox.Text) && !string.IsNullOrWhiteSpace(Address_textBox.Text))
            {
                if (!Regex.IsMatch(Email_textBox.Text, @"^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"))
                {
                    MessageBox.Show("Недействительная электронная почта, проверьте пожалуйста ее написание",
                        "Предупреждение",MessageBoxButton.OK,MessageBoxImage.Hand);
                    Email_textBox.Focus();
                }
                else
                {
                    if(Phone_textBox.Text.Length <11)
                    {
                        MessageBox.Show("Недействительный телефон, проверьте пожалуйста его написание",
                       "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Hand);
                        Phone_textBox.Focus();
                    }

                    else
                    {
                        MessageBox.Show("Вы успешно создали заказ\nВы можете следить за его статус по вашей почте",
                            "Успешный заказ", MessageBoxButton.OK, MessageBoxImage.Information);

                        OrderReadyNot.Email = Email_textBox.Text;
                        OrderReadyNot.Phone = Phone_textBox.Text;
                        OrderReadyNot.Address = Address_textBox.Text;
                        OrderReadyNot.DateOrder = Convert.ToDateTime(DateTime.Now.ToShortDateString());
                        OrderReadyNot.Accepted = false;

                        Dbcontext.Order.Add(OrderReadyNot);
                        Dbcontext.SaveChanges();
                        isChecked = true;
                        Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("Пожалуйста запольните все поля", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }

        private void Back_button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
