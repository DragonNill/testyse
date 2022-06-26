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
    /// Логика взаимодействия для UpdateOrCreateEmployeerWindow.xaml
    /// </summary>
    public partial class UpdateOrCreateEmployeerWindow : Window
    {
        public CreatePcForThePeopl_dbEntities DbContext;
        public Employeer EmployeerID;

        public UpdateOrCreateEmployeerWindow(Employeer employeer)
        {
            InitializeComponent();

            //DbContext = EmployeersShowWindow.DbContext;
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            EmployeerID = employeer;
        }



        private void Back_button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (EmployeerID != null)
            {
                Name_textBox.Text = EmployeerID.Name;
                Second_textBox.Text = EmployeerID.SecondName;
                LastName_textBox.Text = EmployeerID.LastName;
                Login_textBox.Text = EmployeerID.Login;
                Password_textBox.Text = EmployeerID.Password;

                UpdateOrCreate_button.Content = "Редактировать";
                Delete_button.Visibility = Visibility.Visible;
            }

        }

        private void UpdateOrCreate_button_Click(object sender, RoutedEventArgs e)
        {
            if (UpdateOrCreate_button.Content.ToString() == "Редактировать")
            {
                if (!string.IsNullOrWhiteSpace(Name_textBox.Text) && !string.IsNullOrWhiteSpace(Second_textBox.Text)
                    && !string.IsNullOrWhiteSpace(Login_textBox.Text) && !string.IsNullOrWhiteSpace(Password_textBox.Text))
                {

                    EmployeerID.Name = Name_textBox.Text;
                    EmployeerID.SecondName = Second_textBox.Text;
                    if (!string.IsNullOrWhiteSpace(LastName_textBox.Text))
                    {
                        EmployeerID.LastName = LastName_textBox.Text;
                    }
                    EmployeerID.Login = Login_textBox.Text;
                    EmployeerID.Password = Password_textBox.Text;
                    DbContext.SaveChanges();
                    Close();
                }
                else
                {
                    MessageBox.Show("Пожалуйста заполните поля пожалуйста", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Hand);
                }
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(Name_textBox.Text) && !string.IsNullOrWhiteSpace(Second_textBox.Text)
                   && !string.IsNullOrWhiteSpace(Login_textBox.Text) && !string.IsNullOrWhiteSpace(Password_textBox.Text))
                {

                    Employeer employeer = new Employeer();
                    employeer.Name = Name_textBox.Text;
                    employeer.SecondName = Second_textBox.Text;
                    if (!string.IsNullOrWhiteSpace(LastName_textBox.Text))
                    {
                        employeer.LastName = LastName_textBox.Text;
                    }
                    employeer.Login = Login_textBox.Text;
                    employeer.Password = Password_textBox.Text;
                    DbContext.Employeer.Add(employeer);
                    DbContext.SaveChanges();
                    Close();
                }
                else
                {
                    MessageBox.Show("Пожалуйста заполните поля пожалуйста", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Hand);
                }
            }
        }

        private void Delete_button_Click(object sender, RoutedEventArgs e)
        {
            DbContext.Employeer.Remove(EmployeerID);
            DbContext.SaveChanges();
            Close();
        }
    }
}
