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
    /// Логика взаимодействия для EmployeersShowWindow.xaml
    /// </summary>
    public partial class EmployeersShowWindow : Window
    {
        public static CreatePcForThePeopl_dbEntities DbContext;
        public Employeer EmployeerID;

        public EmployeersShowWindow()
        {
            InitializeComponent();

            DbContext = CreatePcForThePeopl_dbEntities.DBContext;

            UpdateListView();
        }

        private void Back_button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void UpdateListView()
        {
            Emploeeyrs_listView.Items.Clear();

            List<Employeer> employeers = new List<Employeer>();

            employeers = DbContext.Employeer.ToList();

            foreach (Employeer employeer in employeers)
            {
                Emploeeyrs_listView.Items.Add($"{employeer.Name.Split(' ')[0]} {employeer.SecondName.Split(' ')[0]} {employeer.LastName.Split(' ')[0]} \nЛогин: {employeer.Login}");
            }
        }

        private void CreateEmployeer_button_Click(object sender, RoutedEventArgs e)
        {
            Window nextWindow = new UpdateOrCreateEmployeerWindow(null);
            nextWindow.ShowDialog();
            UpdateListView();
        }

        private void Emploeeyrs_listView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (Emploeeyrs_listView.SelectedItem != null)
            {
                Employeer employeer = DbContext.Employeer.Where(l 
                    => Emploeeyrs_listView.SelectedItem.ToString().Contains(l.Name) 
                && Emploeeyrs_listView.SelectedItem.ToString().Contains(l.SecondName)).FirstOrDefault();

                new UpdateOrCreateEmployeerWindow(employeer).ShowDialog();
            }
            UpdateListView();
        }
    }
}
