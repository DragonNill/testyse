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
using System.Windows.Shapes;

namespace PcSborka.Views.Windows
{
    /// <summary>
    /// Логика взаимодействия для Authorization.xaml
    /// </summary>
    public partial class Authorization : Window
    {
        public CreatePcForThePeopl_dbEntities context;
        public Employeer user;

        public Authorization()
        {
            InitializeComponent();
            context = new CreatePcForThePeopl_dbEntities();
        }

        private void Back_button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Enter_button_Click(object sender, RoutedEventArgs e)
        {
            if (ShowPassword_checkBox.IsChecked != true)
            {
                Password_textBox.Text = Password_passwordBox.Password;
            }
            if (!string.IsNullOrWhiteSpace(Login_textBox.Text) && !string.IsNullOrWhiteSpace(Password_textBox.Text))
            {

                user = CreatePcForThePeopl_dbEntities.DBContext.Employeer.FirstOrDefault(l => l.Login.Trim() == Login_textBox.Text.Trim() 
                && l.Password.Trim() == Password_textBox.Text.Trim());

                //List<Employeer> employeers = CreatePcForThePeopl_dbEntities.DBContext.Employeer.ToList();
                //foreach (var item in employeers)
                //{
                //    string text = $"Login: {item.Login.Trim()} Password: {item.Password.Trim()}";
                //    MessageBox.Show(text);
                //}

                if (user == null)
                {
                    MessageBox.Show("Неправильный логин или пароль", "Предупреждение"
                        , MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                else
                {
                    Hide();
                    Window nextWindow = new MenuForTheEmployeer(user);
                    nextWindow.ShowDialog();
                    Close();
                }
            }

            else
            {
                MessageBox.Show("Запольните поля пожалуйста", "Предупреждение"
                        ,MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ShowPassword_checkBox_Checked(object sender, RoutedEventArgs e)
        {
            Password_textBox.Text = Password_passwordBox.Password;

            Password_passwordBox.Visibility = Visibility.Hidden;
            Password_textBox.Visibility = Visibility.Visible;
        }

        private void ShowPassword_checkBox_Unchecked(object sender, RoutedEventArgs e)
        {
            Password_passwordBox.Password = Password_textBox.Text;

            Password_passwordBox.Visibility = Visibility.Visible;
            Password_textBox.Visibility = Visibility.Hidden;
        }
    }
}
