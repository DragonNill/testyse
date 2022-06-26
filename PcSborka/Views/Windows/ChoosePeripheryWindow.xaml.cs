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
    /// Логика взаимодействия для ChoosePeripheryWindow.xaml
    /// </summary>
    public partial class ChoosePeripheryWindow : Window
    {
        CreatePcForThePeopl_dbEntities Dbcontext;
        public static int countPeriphery = 0;
        public static Window currentWindow { get; set; }
        public bool isChecked = false;
        public double sumPeriphery = 0;
        public static bool isFinish = false;

        public ChoosePeripheryWindow()
        {
            InitializeComponent();

            Dbcontext = CreatePcForThePeopl_dbEntities.DBContext;
        }

        public ChoosePeripheryWindow(int currentItem)
        {
            InitializeComponent();

            Dbcontext = CreatePcForThePeopl_dbEntities.DBContext;

            isChecked = true;

            CreatePeriphery_button.Content = "Выбрать";
            CreatePeriphery_button.IsEnabled = true;
            
            if(isFinish)
            {
                ChooseKeyboard_button.IsEnabled = false;

                ChooseMonitor_button.IsEnabled = false;

                ChooseMouse_button.IsEnabled = false;

                CreatePeriphery_button.Visibility = Visibility.Hidden;
            }

            if (ChooseCurrentItemShow.PeripheryReadyNot.Keyboard != null)
            {

                Keyboard_listView.Items.Clear();

                Keyboard_listView.Items.Add(new CurrentItemControl(ChooseCurrentItemShow.PeripheryReadyNot.Keyboard, 10));
                ChooseKeyboard_button.IsEnabled = false;
            }

            if (ChooseCurrentItemShow.PeripheryReadyNot.Monitor != null)
            {
                Monitor_listView.Items.Clear();

                Monitor_listView.Items.Add(new CurrentItemControl(ChooseCurrentItemShow.PeripheryReadyNot.Monitor, 8));

                ChooseMonitor_button.IsEnabled = false;
            }

            if (ChooseCurrentItemShow.PeripheryReadyNot.Mouse != null)
            {
                Mouse_listView.Items.Clear();

                Mouse_listView.Items.Add(new CurrentItemControl(ChooseCurrentItemShow.PeripheryReadyNot.Mouse, 9));

                ChooseMouse_button.IsEnabled = false; 
            }

        }

        private void ChooseMonitor_button_Click(object sender, RoutedEventArgs e)
        {
            CheckCount();
            Window nextWindow = new ChooseCurrentItemShow(8);
            nextWindow.ShowDialog();


            if (ChooseCurrentItemShow.PeripheryReadyNot.Monitor != null)
            {
                ChooseMonitor_button.IsEnabled = false;

                Monitor_listView.Items.Clear();

                Monitor_listView.Items.Add(new CurrentItemControl(ChooseCurrentItemShow.PeripheryReadyNot.Monitor, 8));

                sumPeriphery += ChooseCurrentItemShow.PeripheryReadyNot.Monitor.Cost;

                CheckCountinThree();

                CountIsExist();
            }
        }

        private void ChooseMouse_button_Click(object sender, RoutedEventArgs e)
        {
            CheckCount();
            Window nextWindow = new ChooseCurrentItemShow(9);
            nextWindow.ShowDialog();


            if (ChooseCurrentItemShow.PeripheryReadyNot.Mouse != null)
            {
                ChooseMouse_button.IsEnabled = false;

                Mouse_listView.Items.Clear();

                Mouse_listView.Items.Add(new CurrentItemControl(ChooseCurrentItemShow.PeripheryReadyNot.Mouse, 9));

                sumPeriphery += ChooseCurrentItemShow.PeripheryReadyNot.Mouse.Cost;

                CheckCountinThree();

                CountIsExist();
            }
        }

        private void ChooseKeyboard_button_Click(object sender, RoutedEventArgs e)
        {
            CheckCount();
            Window nextWindow = new ChooseCurrentItemShow(10);
            nextWindow.ShowDialog();

            if (ChooseCurrentItemShow.PeripheryReadyNot.Keyboard != null)
            {
                ChooseKeyboard_button.IsEnabled = false;

                Keyboard_listView.Items.Clear();

                Keyboard_listView.Items.Add(new CurrentItemControl(ChooseCurrentItemShow.PeripheryReadyNot.Keyboard, 10));

                sumPeriphery += ChooseCurrentItemShow.PeripheryReadyNot.Keyboard.Cost;

                CheckCountinThree();

                CountIsExist();  
            }
        }

        private void Back_button_Click(object sender, RoutedEventArgs e)
        {
            ChooseCurrentItemShow.PeripheryReadyNot = new Periphery();
            countPeriphery = 0;
            Close();
        }

        private void CountIsExist()
        {
            if(countPeriphery != 0)
            {
                CreatePeriphery_button.Visibility = Visibility.Visible;
            }
        }

        private void CheckCount()
        {
            switch (countPeriphery)
            {
                case 0:
                    countPeriphery++;
                    break;
                case 1:
                    countPeriphery++;
                    break;
                case 2:
                    countPeriphery++;
                    break;
            }
        }

        private void CheckCountinThree()
        {
            if (countPeriphery == 3)
            {
                

                
            }
        }

        private void CreatePeriphery_button_Click(object sender, RoutedEventArgs e)
        {
            if (!isChecked)
            {
                FinalStageOfOrder.OrderReadyNot.Periphery = ChooseCurrentItemShow.PeripheryReadyNot;
                ChooseCurrentItemShow.PeripheryReadyNot.SumPeriphery = sumPeriphery;
                Dbcontext.Periphery.Add(ChooseCurrentItemShow.PeripheryReadyNot);
                Dbcontext.SaveChanges();
                Close();
            }
            else
            {
                FinalStageOfOrder.OrderReadyNot.Periphery = ChooseCurrentItemShow.PeripheryReadyNot;
                currentWindow.Close();
                currentWindow = null;
                Close();
            }
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            ChooseCurrentItemShow.PeripheryReadyNot.isBuildPeriphery = true;
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            ChooseCurrentItemShow.PeripheryReadyNot.isBuildPeriphery = false;
        }
    }
}
