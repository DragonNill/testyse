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
    /// Логика взаимодействия для InputEmailWindow.xaml
    /// </summary>
    public partial class InputEmailWindow : Window
    {
        public CreatePcForThePeopl_dbEntities Dbcontext;
        public Order order;
        public string Email;

        public InputEmailWindow()
        {
            InitializeComponent();

            Dbcontext = CreatePcForThePeopl_dbEntities.DBContext;
        }

        private void NextWindow_button_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(Email_textBox.Text))
            {
                if (Regex.IsMatch(Email_textBox.Text, @"^[a-zA-Z][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$"))
                {
                    Email = Email_textBox.Text;
                    if(Dbcontext.Order.Any(i =>i.Email ==Email))
                    {
                        MessageBox.Show("Есть");
                    }
                    else
                    {
                        MessageBox.Show("nety");
                    }
                }
                else
                {
                    MessageBox.Show("Недействительная электронная почта,\nпожалуйста провертьте почту", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Hand);
                    Email_textBox.Focus();
                }
            }
            else
            {
                MessageBox.Show("Запольните поля","Предупреждение",MessageBoxButton.OK,MessageBoxImage.Hand);
                Email_textBox.Focus();
            }

        }
    }
}
