using PcSborka.Entity;
using System.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    /// Логика взаимодействия для ProductInfoWindow.xaml
    /// </summary>
    public partial class ProductInfoWindow : Window
    {
        CreatePcForThePeopl_dbEntities DbContext;
        public CPU CPUID { get; set; }
        public MotherBoard MotherBoardID { get; set; }
        public Case CaseID { get; set; }
        public GPU GPUID { get; set; }
        public Cooler CoolerID { get; set; }
        public RAM RAMID { get; set; }
        public PowerSupply PowerSupplyID { get; set; }
        public Monitor MonitorID { get; set; }
        public Entity.Mouse MouseID { get; set; }
        public Entity.Keyboard KeyboardID { get; set; }
        public int CurrentItem;
        public static bool isCheckys = false;
        public static bool checkys = false;
        public string[] splitTextBox, splitTextBox2, splitTextBox3, splitTextBox4;

        public static Window creatingForm { get; set; }

        #region Constructs
        public ProductInfoWindow(CPU cpu, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            CPUID = cpu;
            CurrentItem = currentItem;
        }

        public ProductInfoWindow(MotherBoard motherBoard, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            MotherBoardID = motherBoard;
            CurrentItem = currentItem;
        }

        public ProductInfoWindow(Case @case, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            CaseID = @case;
            CurrentItem = currentItem;
        }

        public ProductInfoWindow(GPU gPU, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            GPUID = gPU;
            CurrentItem = currentItem;
        }

        public ProductInfoWindow(Cooler cooler, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            CoolerID = cooler;
            CurrentItem = currentItem;
        }

        public ProductInfoWindow(RAM rAM, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            RAMID = rAM;
            CurrentItem = currentItem;
        }

        public ProductInfoWindow(PowerSupply powerSupply, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            PowerSupplyID = powerSupply;
            CurrentItem = currentItem;
        }

        public ProductInfoWindow(Monitor monitor, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            MonitorID = monitor;
            CurrentItem = currentItem;
        }

        public ProductInfoWindow(Entity.Mouse mouse, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            MouseID = mouse;
            CurrentItem = currentItem;
        }

        public ProductInfoWindow(Entity.Keyboard keyboard, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            KeyboardID = keyboard;
            CurrentItem = currentItem;
        }
        #endregion 

        private void Back_button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (isCheckys)
            {
                AddToChoose_button.Content = "Удалить";
                Update_button.Visibility = Visibility.Visible;

                Cost_textBox.IsEnabled = true;
                Charastericts_textBox.IsEnabled = true;
                Form_Factor_textBox.IsEnabled = true;
                Socket_textBox.IsEnabled = true;
                Title_textBox.IsEnabled = true;
                Form_Factor_textBox.IsEnabled = true;

                Socket_textBox.Visibility = Visibility.Visible;
                TypeMemory_textBox.Visibility = Visibility.Visible;
                Form_Factor_textBox.Visibility = Visibility.Visible;
            }
            if(checkys)
            {
                AddToChoose_button.Content = "Добавить новый товар";

                Cost_textBox.IsEnabled = true;
                Charastericts_textBox.IsEnabled = true;
                Form_Factor_textBox.IsEnabled = true;
                Socket_textBox.IsEnabled = true;
                Title_textBox.IsEnabled = true;
                Form_Factor_textBox.IsEnabled = true;

                Socket_textBox.Visibility = Visibility.Visible;
                TypeMemory_textBox.Visibility = Visibility.Visible;
                Form_Factor_textBox.Visibility = Visibility.Visible;
            }
            CheckItem(CurrentItem);
        }

        private void CheckItem(int currentItem)
        {
            if (checkys)
            {
                Cost_textBox.Width = 300;
                Title_textBox.Width = 300;
                Charastericts_textBox.Width = 300;
                Charastericts_textBox.Height = 150;
                switch (currentItem)
                {
                    case 1:
                        Cost_textBox.Text = $"Цена: ";
                        Socket_textBox.Text = $"Сокет: ";
                        Socket_textBox.Visibility = Visibility.Visible;
                        break;
                    case 2:
                        Cost_textBox.Text = $"Цена: ";
                        Socket_textBox.Text = $"Сокет: ";
                        Socket_textBox.Visibility = Visibility.Visible;
                        TypeMemory_textBox.Visibility = Visibility.Visible;
                        TypeMemory_textBox.Text = $"Тип памяти: ";
                        Form_Factor_textBox.Visibility = Visibility.Visible;
                        Form_Factor_textBox.Text = $"Форм-фактор: ";
                        break;
                    case 3:
                        Cost_textBox.Text = $"Цена: ";
                        Form_Factor_textBox.Visibility = Visibility.Visible;
                        Form_Factor_textBox.Text = $"Форм-фактор: ";
                        break;
                    case 4:
                        Cost_textBox.Text = $"Цена: ";
                        break;
                    case 5:
                        Cost_textBox.Text = $"Цена: ";
                        break;
                    case 6:
                        Cost_textBox.Text = $"Цена: ";
                        TypeMemory_textBox.Visibility = Visibility.Visible;
                        TypeMemory_textBox.Text = $"Тип памяти: ";
                        break;
                    case 7:
                        Cost_textBox.Text = $"Цена: ";
                        break;
                    case 8:
                        Cost_textBox.Text = $"Цена: ";
                        break;
                    case 9:
                        Cost_textBox.Text = $"Цена: ";
                        Charastericts_textBox.Visibility = Visibility.Hidden;
                        break;
                    case 10:
                        Cost_textBox.Text = $"Цена: ";
                        Charastericts_textBox.Visibility = Visibility.Hidden;
                        break;
                }
            }
            else
            {
                switch (currentItem)
                {
                    case 1:
                        Title_textBox.Text = CPUID.Name;
                        Charastericts_textBox.Text = CPUID.Сharacteristics;
                        Cost_textBox.Text = $"Цена: {CPUID.Cost} руб.";
                        Socket_textBox.Visibility = Visibility.Visible;
                        Socket_textBox.Text = $"Сокет: {CPUID.Socket}";
                        break;
                    case 2:
                        Title_textBox.Text = MotherBoardID.Name;
                        Cost_textBox.Text = $"Цена: {MotherBoardID.Cost} руб.";
                        Charastericts_textBox.Text = MotherBoardID.Сharacteristics;
                        Socket_textBox.Text = $"Сокет: {MotherBoardID.Socket}";
                        Socket_textBox.Visibility = Visibility.Visible;
                        TypeMemory_textBox.Visibility = Visibility.Visible;
                        TypeMemory_textBox.Text = $"Тип памяти: {MotherBoardID.TypeMotherBoardMemory}";
                        Form_Factor_textBox.Visibility = Visibility.Visible;
                        Form_Factor_textBox.Text = $"Форм-фактор: {MotherBoardID.Form_Factor}";
                        break;
                    case 3:
                        Title_textBox.Text = CaseID.Name;
                        Cost_textBox.Text = $"Цена: {CaseID.Cost} руб.";
                        Charastericts_textBox.Text = CaseID.Сharacteristics;
                        Form_Factor_textBox.Visibility = Visibility.Visible;
                        Form_Factor_textBox.Text = $"Форм-фактор {CaseID.Form_Factor}";
                        break;
                    case 4:
                        Title_textBox.Text = GPUID.Name;
                        Charastericts_textBox.Text = GPUID.Сharacteristics;
                        Cost_textBox.Text = $"Цена: {GPUID.Cost} руб.";
                        break;
                    case 5:
                        Title_textBox.Text = CoolerID.Name;
                        Charastericts_textBox.Text = CoolerID.Сharacteristics;
                        Cost_textBox.Text = $"Цена: {CoolerID.Cost} руб.";
                        break;
                    case 6:
                        Title_textBox.Text = RAMID.Name;
                        Cost_textBox.Text = $"Цена: {RAMID.Cost} руб.";
                        Charastericts_textBox.Text = RAMID.Сharacteristics;
                        TypeMemory_textBox.Visibility = Visibility.Visible;
                        TypeMemory_textBox.Text = $"Тип памяти: {RAMID.TypeMemory}";
                        break;
                    case 7:
                        Title_textBox.Text = PowerSupplyID.Name;
                        Charastericts_textBox.Text = PowerSupplyID.Сharacteristics;
                        Cost_textBox.Text = $"Цена: {PowerSupplyID.Cost} руб.";
                        break;
                    case 8:
                        Title_textBox.Text = MonitorID.Name;
                        Charastericts_textBox.Text = MonitorID.Сharacteristics;
                        Cost_textBox.Text = $"Цена: {MonitorID.Cost} руб.";
                        break;
                    case 9:
                        Title_textBox.Text = MouseID.Name;
                        Cost_textBox.Text = $"Цена: {MouseID.Cost} руб.";
                        Charastericts_textBox.Visibility = Visibility.Hidden;
                        break;
                    case 10:
                        Title_textBox.Text = KeyboardID.Name;
                        Cost_textBox.Text = $"Цена: {KeyboardID.Cost} руб.";
                        Charastericts_textBox.Visibility = Visibility.Hidden;

                        break;
                }
            }
        }

        private void AddToChoose_button_Click(object sender, RoutedEventArgs e)
        {
            if (!checkys)
            {
                if (!isCheckys)
                {
                    switch (CurrentItem)
                    {
                        case 1:
                            ChooseCurrentItemShow.ComputerReadyNot.CPU = CPUID;
                            creatingForm.Close();
                            Close();
                            break;
                        case 2:
                            ChooseCurrentItemShow.ComputerReadyNot.MotherBoard = MotherBoardID;
                            creatingForm.Close();
                            Close();
                            break;
                        case 3:
                            ChooseCurrentItemShow.ComputerReadyNot.Case = CaseID;
                            creatingForm.Close();
                            Close();
                            break;
                        case 4:
                            ChooseCurrentItemShow.ComputerReadyNot.GPU = GPUID;
                            creatingForm.Close();
                            Close();
                            break;
                        case 5:
                            ChooseCurrentItemShow.ComputerReadyNot.Cooler = CoolerID;
                            creatingForm.Close();
                            Close();
                            break;
                        case 6:
                            ChooseCurrentItemShow.ComputerReadyNot.RAM = RAMID;
                            creatingForm.Close();
                            Close();
                            break;
                        case 7:
                            ChooseCurrentItemShow.ComputerReadyNot.PowerSupply = PowerSupplyID;
                            creatingForm.Close();
                            Close();
                            break;
                        case 8:
                            ChooseCurrentItemShow.PeripheryReadyNot.Monitor = MonitorID;
                            creatingForm.Close();
                            Close();
                            break;
                        case 9:
                            ChooseCurrentItemShow.PeripheryReadyNot.Mouse = MouseID;
                            creatingForm.Close();
                            Close();
                            break;
                        case 10:
                            ChooseCurrentItemShow.PeripheryReadyNot.Keyboard = KeyboardID;
                            creatingForm.Close();
                            Close();
                            break;
                    }
                }
                else
                {
                    switch (CurrentItem)
                    {
                        case 1:
                            DbContext.CPU.Remove(CPUID);
                            DbContext.SaveChanges();
                            MessageBox.Show("Вы успешно удалили товар", "Уведомление");
                            Close();
                            break;
                        case 2:
                            DbContext.MotherBoard.Remove(MotherBoardID);
                            DbContext.SaveChanges();
                            MessageBox.Show("Вы успешно удалили товар", "Уведомление");
                            Close();
                            break;
                        case 3:
                            DbContext.Case.Remove(CaseID);
                            DbContext.SaveChanges();
                            MessageBox.Show("Вы успешно удалили товар", "Уведомление");
                            Close();
                            break;
                        case 4:
                            DbContext.GPU.Remove(GPUID);
                            DbContext.SaveChanges();
                            MessageBox.Show("Вы успешно удалили товар", "Уведомление");
                            Close();
                            break;
                        case 5:
                            DbContext.Cooler.Remove(CoolerID);
                            DbContext.SaveChanges();
                            MessageBox.Show("Вы успешно удалили товар", "Уведомление");
                            Close();
                            break;
                        case 6:
                            DbContext.RAM.Remove(RAMID);
                            DbContext.SaveChanges();
                            MessageBox.Show("Вы успешно удалили товар", "Уведомление");
                            Close();
                            break;
                        case 7:
                            DbContext.PowerSupply.Remove(PowerSupplyID);
                            DbContext.SaveChanges();
                            MessageBox.Show("Вы успешно удалили товар", "Уведомление");
                            Close();
                            break;
                        case 8:
                            DbContext.Monitor.Remove(MonitorID);
                            DbContext.SaveChanges();
                            MessageBox.Show("Вы успешно удалили товар", "Уведомление");
                            Close();
                            break;
                        case 9:
                            DbContext.Mouse.Remove(MouseID);
                            DbContext.SaveChanges();
                            MessageBox.Show("Вы успешно удалили товар", "Уведомление");
                            Close();
                            break;
                        case 10:
                            DbContext.Keyboard.Remove(KeyboardID);
                            DbContext.SaveChanges();
                            MessageBox.Show("Вы успешно удалили товар", "Уведомление");
                            Close();
                            break;
                    }
                }
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(Cost_textBox.Text) && !string.IsNullOrWhiteSpace(Title_textBox.Text))
                {
                    switch (CurrentItem)
                    {
                        case 1:
                            CPUID.Name = Title_textBox.Text;
                            splitTextBox = Cost_textBox.Text.Split(' ');
                            CPUID.Cost = Convert.ToDouble(splitTextBox[1]);
                            splitTextBox2 = Socket_textBox.Text.Split(' ');
                            CPUID.Socket = splitTextBox2[1];
                            CPUID.Сharacteristics = Charastericts_textBox.Text;
                            DbContext.CPU.Add(CPUID);
                            DbContext.SaveChanges();
                            Close();
                            break;
                        case 2:
                            MotherBoardID.Name = Title_textBox.Text;
                            splitTextBox = Cost_textBox.Text.Split(' ');
                            MotherBoardID.Cost = Convert.ToDouble(splitTextBox[1]);
                            splitTextBox2 = Socket_textBox.Text.Split(' ');
                            MotherBoardID.Socket = splitTextBox2[1];
                            MotherBoardID.Сharacteristics = Charastericts_textBox.Text;
                            splitTextBox3 = Form_Factor_textBox.Text.Split(' ');
                            MotherBoardID.Form_Factor = splitTextBox3[1];
                            splitTextBox4 = TypeMemory_textBox.Text.Split(' ');
                            MotherBoardID.TypeMotherBoardMemory = splitTextBox4[2];
                            DbContext.MotherBoard.Add(MotherBoardID);
                            DbContext.SaveChanges();
                            Close();
                            break;
                        case 3:
                            CaseID.Name = Title_textBox.Text;
                            splitTextBox = Cost_textBox.Text.Split(' ');
                            CaseID.Cost = Convert.ToDouble(splitTextBox[1]);
                            CaseID.Сharacteristics = Charastericts_textBox.Text;
                            splitTextBox3 = Form_Factor_textBox.Text.Split(' ');
                            CaseID.Form_Factor = splitTextBox3[1];
                            DbContext.Case.Add(CaseID);
                            DbContext.SaveChanges();
                            Close();
                            break;
                        case 4:
                            GPUID.Name = Title_textBox.Text;
                            splitTextBox = Cost_textBox.Text.Split(' ');
                            GPUID.Cost = Convert.ToDouble(splitTextBox[1]);
                            GPUID.Сharacteristics = Charastericts_textBox.Text;
                            DbContext.GPU.Add(GPUID);
                            DbContext.SaveChanges();
                            Close();
                            break;
                        case 5:
                            CoolerID.Name = Title_textBox.Text;
                            splitTextBox = Cost_textBox.Text.Split(' ');
                            CoolerID.Cost = Convert.ToDouble(splitTextBox[1]);
                            CoolerID.Сharacteristics = Charastericts_textBox.Text;
                            DbContext.Cooler.Add(CoolerID);
                            DbContext.SaveChanges();
                            Close();
                            break;
                        case 6:
                            RAMID.Name = Title_textBox.Text;
                            splitTextBox = Cost_textBox.Text.Split(' ');
                            RAMID.Cost = Convert.ToDouble(splitTextBox[1]);
                            RAMID.Сharacteristics = Charastericts_textBox.Text;
                            RAMID.TypeMemory = TypeMemory_textBox.Text;
                            DbContext.RAM.Add(RAMID);
                            DbContext.SaveChanges();
                            Close();
                            break;
                        case 7:
                            PowerSupplyID.Name = Title_textBox.Text;
                            splitTextBox = Cost_textBox.Text.Split(' ');
                            PowerSupplyID.Cost = Convert.ToDouble(splitTextBox[1]);
                            PowerSupplyID.Сharacteristics = Charastericts_textBox.Text;
                            DbContext.PowerSupply.Add(PowerSupplyID);
                            DbContext.SaveChanges();
                            Close();
                            break;
                        case 8:
                            MonitorID.Name = Title_textBox.Text;
                            splitTextBox = Cost_textBox.Text.Split(' ');
                            MonitorID.Cost = Convert.ToDouble(splitTextBox[1]);
                            MonitorID.Сharacteristics = Charastericts_textBox.Text;
                            DbContext.Monitor.Add(MonitorID);
                            DbContext.SaveChanges();
                            Close();
                            break;
                        case 9:
                            MouseID.Name = Title_textBox.Text;
                            splitTextBox = Cost_textBox.Text.Split(' ');
                            MouseID.Cost = Convert.ToDouble(splitTextBox[1]);
                            DbContext.Mouse.Add(MouseID);
                            DbContext.SaveChanges();
                            Close();
                            break;
                        case 10:
                            KeyboardID.Name = Title_textBox.Text;
                            splitTextBox = Cost_textBox.Text.Split(' ');
                            KeyboardID.Cost = Convert.ToDouble(splitTextBox[1]);
                            DbContext.Keyboard.Add(KeyboardID);
                            DbContext.SaveChanges();
                            Close();
                            break;
                    }
                }
                else
                {
                    MessageBox.Show("Запольните все поля");
                }
                }
        }

        private void Update_button_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(Title_textBox.Text) && !string.IsNullOrWhiteSpace(Cost_textBox.Text))
            {
                switch (CurrentItem)
                {
                    case 1:
                        CPUID.Name = Title_textBox.Text;
                        splitTextBox = Cost_textBox.Text.Split(' ');
                        CPUID.Cost = Convert.ToDouble(splitTextBox[1]);
                        splitTextBox2 = Socket_textBox.Text.Split(' ');
                        CPUID.Socket = splitTextBox2[1];
                        CPUID.Сharacteristics = Charastericts_textBox.Text;
                        DbContext.SaveChanges();
                        Close();
                        break;
                    case 2:
                        MotherBoardID.Name = Title_textBox.Text;
                        splitTextBox = Cost_textBox.Text.Split(' ');
                        MotherBoardID.Cost = Convert.ToDouble(splitTextBox[1]);
                        splitTextBox2 = Socket_textBox.Text.Split(' ');
                        MotherBoardID.Socket = splitTextBox2[1];
                        MotherBoardID.Сharacteristics = Charastericts_textBox.Text;
                        splitTextBox3 = Form_Factor_textBox.Text.Split(' ');
                        MotherBoardID.Form_Factor = splitTextBox3[1];
                        splitTextBox4 = TypeMemory_textBox.Text.Split(' ');
                        MotherBoardID.TypeMotherBoardMemory = splitTextBox4[2];
                        DbContext.SaveChanges();
                        Close();
                        break;
                    case 3:
                        CaseID.Name = Title_textBox.Text;
                        splitTextBox = Cost_textBox.Text.Split(' ');
                        CaseID.Cost = Convert.ToDouble(splitTextBox[1]);
                        CaseID.Сharacteristics = Charastericts_textBox.Text;
                        splitTextBox3 = Form_Factor_textBox.Text.Split(' ');
                        CaseID.Form_Factor = splitTextBox3[1];
                        DbContext.SaveChanges();
                        Close();
                        break;
                    case 4:
                        GPUID.Name = Title_textBox.Text;
                        splitTextBox = Cost_textBox.Text.Split(' ');
                        GPUID.Cost = Convert.ToDouble(splitTextBox[1]);
                        GPUID.Сharacteristics = Charastericts_textBox.Text;
                        DbContext.SaveChanges();
                        Close();
                        break;
                    case 5:
                        CoolerID.Name = Title_textBox.Text;
                        splitTextBox = Cost_textBox.Text.Split(' ');
                        CoolerID.Cost = Convert.ToDouble(splitTextBox[1]);
                        CoolerID.Сharacteristics = Charastericts_textBox.Text;
                        DbContext.SaveChanges();
                        Close();
                        break;
                    case 6:
                        RAMID.Name = Title_textBox.Text;
                        splitTextBox = Cost_textBox.Text.Split(' ');
                        RAMID.Cost = Convert.ToDouble(splitTextBox[1]);
                        RAMID.Сharacteristics = Charastericts_textBox.Text;
                        RAMID.TypeMemory = TypeMemory_textBox.Text;
                        DbContext.SaveChanges();
                        Close();
                        break;
                    case 7:
                        PowerSupplyID.Name = Title_textBox.Text;
                        splitTextBox = Cost_textBox.Text.Split(' ');
                        PowerSupplyID.Cost = Convert.ToDouble(splitTextBox[1]);
                        PowerSupplyID.Сharacteristics = Charastericts_textBox.Text;
                        DbContext.SaveChanges();
                        Close();
                        break;
                    case 8:
                        MonitorID.Name = Title_textBox.Text;
                        splitTextBox = Cost_textBox.Text.Split(' ');
                        MonitorID.Cost = Convert.ToDouble(splitTextBox[1]);
                        MonitorID.Сharacteristics = Charastericts_textBox.Text;
                        DbContext.SaveChanges();
                        Close();
                        break;
                    case 9:
                        MouseID.Name = Title_textBox.Text;
                        splitTextBox = Cost_textBox.Text.Split(' ');
                        MouseID.Cost = Convert.ToDouble(splitTextBox[1]);
                        DbContext.SaveChanges();
                        Close();
                        break;
                    case 10:
                        KeyboardID.Name = Title_textBox.Text;
                        splitTextBox = Cost_textBox.Text.Split(' ');
                        KeyboardID.Cost = Convert.ToDouble(splitTextBox[1]);
                        DbContext.SaveChanges();
                        Close();
                        break;
                }
            }
            else
            {
                MessageBox.Show("Запольните поля прежде чем редактировать","Уведомление");
            }
        }

        private void Cost_textBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !(Char.IsDigit(e.Text, 0));
        }

        private void Cost_textBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                e.Handled = true;
            }

        }
    }
}
