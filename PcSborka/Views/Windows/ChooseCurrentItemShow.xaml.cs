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
    /// Логика взаимодействия для ChooseCurrentItemShow.xaml
    /// </summary>
    public partial class ChooseCurrentItemShow : Window
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
        public Computer ComputerID { get; set; }
        public Periphery PeripheryID { get; set; }
        public int CurrentItem;
        public string TypeMotherBoardT, SocketT, Form_factorT;
        public static Computer ComputerReadyNot { get; set; } = new Computer();
        public static Periphery PeripheryReadyNot { get; set; } = new Periphery();
        public bool isCheckys = false;
        public static bool checkys = false;

        public ChooseCurrentItemShow(int currentItem)
        {
            InitializeComponent();

            CurrentItem = currentItem;
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            if (currentItem == 11 || currentItem == 12)
            {
                Find_textBlock.Visibility = Visibility.Hidden;
                Find_textBox.Visibility = Visibility.Hidden;
            }
            CheckItem(currentItem);
        }

        public ChooseCurrentItemShow(int currentItem,string check)
        {
            InitializeComponent();

            CurrentItem = currentItem;
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;

            AddToChoose_button.Content = "Удалить";

            AddNewItem_button.Visibility = Visibility.Visible;
            isCheckys = true;
            CheckItem(currentItem);
        }


        public ChooseCurrentItemShow(int currentItem, string findAtribite, int findAtribiteID)
        {
            InitializeComponent();

            CurrentItem = currentItem;
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;

            switch (findAtribiteID)
            {
                case 1:
                    SocketT = findAtribite;
                    CheckItem(currentItem);
                    break;
                case 2:
                    Form_factorT = findAtribite;
                    CheckItem(currentItem);
                    break;
                case 3:
                    TypeMotherBoardT = findAtribite;
                    CheckItem(currentItem);
                    break;
            }
        }
        private void Back_button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void CheckItem(int currentItem)
        {
            switch (currentItem)
            {
                case 1:
                    this.Title = "Выбор процессора";

                    CurrentItem_listView.Items.Clear();

                    List<CPU> cpuList = new List<CPU>();
                    cpuList = DbContext.CPU.ToList();

                    if (!string.IsNullOrWhiteSpace(Find_textBox.Text.ToLower()))
                    {
                        cpuList = cpuList.Where(x => x.Name.ToLower().Contains(Find_textBox.Text.ToLower())).ToList();
                    }
                    switch (Sorting_comboBox.SelectedIndex)
                    {
                        case 1:
                            cpuList = cpuList.OrderBy(o => o.Cost).ToList();
                            break;
                        case 2:
                            cpuList = cpuList.OrderByDescending(o => o.Cost).ToList();
                            break;
                    }
                    foreach (CPU cpu in cpuList)
                    {
                        CurrentItem_listView.Items.Add(new CurrentItemControl(cpu, currentItem));
                    }
                    break;
                case 2:
                    this.Title = "Выбор материнская плата";

                    CurrentItem_listView.Items.Clear();

                    List<MotherBoard> motherBoards = new List<MotherBoard>();
                    motherBoards = DbContext.MotherBoard.ToList();

                    if (SocketT != null)
                    {
                        motherBoards = motherBoards.Where(l => l.Socket.Contains(SocketT)).ToList();
                    }

                    
                    if (!string.IsNullOrWhiteSpace(Find_textBox.Text.ToLower()))
                    {
                        motherBoards = motherBoards.Where(x => x.Name.ToLower().Contains(Find_textBox.Text.ToLower())).ToList();
                    }
                    switch (Sorting_comboBox.SelectedIndex)
                    {
                        case 1:
                            motherBoards = motherBoards.OrderBy(o => o.Cost).ToList();
                            break;
                        case 2:
                            motherBoards = motherBoards.OrderByDescending(o => o.Cost).ToList();
                            break;
                    }

                    foreach (MotherBoard motherBoard in motherBoards)
                    {
                        CurrentItem_listView.Items.Add(new CurrentItemControl(motherBoard, currentItem));
                    }
                    break;
                case 3:
                    this.Title = "Выбор корпус";

                    CurrentItem_listView.Items.Clear();

                    List<Case> cases = new List<Case>();
                    cases = DbContext.Case.ToList();

                    if (Form_factorT != null)
                    {
                        cases = cases.Where(c => c.Form_Factor.ToLower().Contains(Form_factorT.ToLower())).ToList();
                    }

                    if (!string.IsNullOrWhiteSpace(Find_textBox.Text.ToLower()))
                    {
                        cases = cases.Where(x => x.Name.ToLower().Contains(Find_textBox.Text.ToLower())).ToList();
                    }

                    switch (Sorting_comboBox.SelectedIndex)
                    {
                        case 1:
                            cases = cases.OrderBy(o => o.Cost).ToList();
                            break;
                        case 2:
                            cases = cases.OrderByDescending(o => o.Cost).ToList();
                            break;
                    }

                    foreach (Case c in cases)
                    {
                        CurrentItem_listView.Items.Add(new CurrentItemControl(c, currentItem));
                    }
                    break;
                case 4:
                    this.Title = "Выбор видеокарта";

                    CurrentItem_listView.Items.Clear();

                    List<GPU> gpus = new List<GPU>();
                    gpus = DbContext.GPU.ToList();

                    if (!string.IsNullOrEmpty(Find_textBox.Text.ToLower()))
                    {
                        gpus = gpus.Where(x => x.Name.ToLower().Contains(Find_textBox.Text.ToLower())).ToList();
                    }

                    switch (Sorting_comboBox.SelectedIndex)
                    {
                        case 1:
                            gpus = gpus.OrderBy(o => o.Cost).ToList();
                            break;
                        case 2:
                            gpus = gpus.OrderByDescending(o => o.Cost).ToList();
                            break;
                    }

                    foreach (GPU gpu in gpus)
                    {
                        CurrentItem_listView.Items.Add(new CurrentItemControl(gpu, currentItem));
                    }
                    break;
                case 5:
                    this.Title = "Выбор кулер";

                    CurrentItem_listView.Items.Clear();

                    List<Cooler> coolers = new List<Cooler>();
                    coolers = DbContext.Cooler.ToList();

                    if (!string.IsNullOrEmpty(Find_textBox.Text.ToLower()))
                    {
                        coolers = coolers.Where(x => x.Name.ToLower().Contains(Find_textBox.Text.ToLower())).ToList();
                    }

                    switch (Sorting_comboBox.SelectedIndex)
                    {
                        case 1:
                            coolers = coolers.OrderBy(o => o.Cost).ToList();
                            break;
                        case 2:
                            coolers = coolers.OrderByDescending(o => o.Cost).ToList();
                            break;
                    }

                    foreach (Cooler cooler in coolers)
                    {
                        CurrentItem_listView.Items.Add(new CurrentItemControl(cooler, currentItem));
                    }
                    break;
                case 6:
                    this.Title = "Выбор оперативная память";

                    CurrentItem_listView.Items.Clear();

                    List<RAM> RAMs = new List<RAM>();
                    RAMs = DbContext.RAM.ToList();

                    if (TypeMotherBoardT != null)
                    {
                        RAMs = RAMs.Where(t => t.TypeMemory.ToLower().Contains(TypeMotherBoardT.ToLower())).ToList();
                    }

                    if (!string.IsNullOrEmpty(Find_textBox.Text.ToLower()))
                    {
                        RAMs = RAMs.Where(x => x.Name.ToLower().Contains(Find_textBox.Text.ToLower())).ToList();
                    }

                    switch (Sorting_comboBox.SelectedIndex)
                    {
                        case 1:
                            RAMs = RAMs.OrderBy(o => o.Cost).ToList();
                            break;
                        case 2:
                            RAMs = RAMs.OrderByDescending(o => o.Cost).ToList();
                            break;
                    }

                    foreach (RAM RAM in RAMs)
                    {
                        CurrentItem_listView.Items.Add(new CurrentItemControl(RAM, currentItem));
                    }
                    break;
                case 7:
                    this.Title = "Выбор блока питания";

                    CurrentItem_listView.Items.Clear();

                    List<PowerSupply> PowerSupplys = new List<PowerSupply>();
                    PowerSupplys = DbContext.PowerSupply.ToList();

                    if (!string.IsNullOrEmpty(Find_textBox.Text.ToLower()))
                    {
                        PowerSupplys = PowerSupplys.Where(x => x.Name.ToLower().Contains(Find_textBox.Text.ToLower())).ToList();
                    }

                    switch (Sorting_comboBox.SelectedIndex)
                    {
                        case 1:
                            PowerSupplys = PowerSupplys.OrderBy(o => o.Cost).ToList();
                            break;
                        case 2:
                            PowerSupplys = PowerSupplys.OrderByDescending(o => o.Cost).ToList();
                            break;
                    }

                    foreach (PowerSupply PowerSupply in PowerSupplys)
                    {
                        CurrentItem_listView.Items.Add(new CurrentItemControl(PowerSupply, currentItem));
                    }
                    break;
                case 8:
                    this.Title = "Выбор монитора";

                    CurrentItem_listView.Items.Clear();

                    List<Monitor> monitors = new List<Monitor>();
                    monitors = DbContext.Monitor.ToList();

                    if (!string.IsNullOrEmpty(Find_textBox.Text.ToLower()))
                    {
                        monitors = monitors.Where(x => x.Name.ToLower().Contains(Find_textBox.Text.ToLower())).ToList();
                    }

                    switch (Sorting_comboBox.SelectedIndex)
                    {
                        case 1:
                            monitors = monitors.OrderBy(o => o.Cost).ToList();
                            break;
                        case 2:
                            monitors = monitors.OrderByDescending(o => o.Cost).ToList();
                            break;
                    }

                    foreach (Monitor monitor in monitors)
                    {
                        CurrentItem_listView.Items.Add(new CurrentItemControl(monitor, currentItem));
                    }
                    break;
                case 9:
                    this.Title = "Выбор мышки";

                    CurrentItem_listView.Items.Clear();

                    List<Entity.Mouse> mice = new List<Entity.Mouse>();
                    mice = DbContext.Mouse.ToList();

                    if (!string.IsNullOrEmpty(Find_textBox.Text.ToLower()))
                    {
                        mice = mice.Where(x => x.Name.ToLower().Contains(Find_textBox.Text.ToLower())).ToList();
                    }

                    switch (Sorting_comboBox.SelectedIndex)
                    {
                        case 1:
                            mice = mice.OrderBy(o => o.Cost).ToList();
                            break;
                        case 2:
                            mice = mice.OrderByDescending(o => o.Cost).ToList();
                            break;
                    }

                    foreach (Entity.Mouse mouse in mice)
                    {
                        CurrentItem_listView.Items.Add(new CurrentItemControl(mouse, currentItem));
                    }
                    break;
                case 10:
                    this.Title = "Выбор клавиатуры";

                    CurrentItem_listView.Items.Clear();

                    List<Entity.Keyboard> keyboards = new List<Entity.Keyboard>();
                    keyboards = DbContext.Keyboard.ToList();

                    if (!string.IsNullOrEmpty(Find_textBox.Text.ToLower()))
                    {
                        keyboards = keyboards.Where(x => x.Name.ToLower().Contains(Find_textBox.Text.ToLower())).ToList();
                    }

                    switch (Sorting_comboBox.SelectedIndex)
                    {
                        case 1:
                            keyboards = keyboards.OrderBy(o => o.Cost).ToList();
                            break;
                        case 2:
                            keyboards = keyboards.OrderByDescending(o => o.Cost).ToList();
                            break;
                    }

                    foreach (Entity.Keyboard keyboard in keyboards)
                    {
                        CurrentItem_listView.Items.Add(new CurrentItemControl(keyboard, currentItem));
                    }
                    break;
                case 11:
                    this.Title = "Выбор компьютера";

                    CurrentItem_listView.Items.Clear();

                    List<Computer> computers = new List<Computer>();

                        computers = DbContext.Computer.Where(w => w.PowerSupply != null).ToList();
                        // w.CPU != null || w.MotherBoard != null || w.Case != null || w.GPU != null || w.Cooler != null || w.RAM != null || 
                    switch (Sorting_comboBox.SelectedIndex)
                    {
                        case 1:
                            computers = computers.OrderBy(o => o.SumComponents).ToList();
                            break;
                        case 2:
                            computers = computers.OrderByDescending(o => o.SumComponents).ToList();
                            break;
                    }

                    foreach (Computer computer in computers)
                    {
                        CurrentItem_listView.Items.Add(new CurrentItemControl(computer, currentItem));
                    }
                    break;
                case 12:
                    this.Title = "Выбор периферий";

                    AddNewItem_button.Content = "Добавить новую сборку";

                    CurrentItem_listView.Items.Clear();

                    List<Periphery> peripheries = new List<Periphery>();

                    peripheries = DbContext.Periphery.ToList();

                    switch (Sorting_comboBox.SelectedIndex)
                    {
                        case 1:
                            peripheries = peripheries.OrderBy(o => o.SumPeriphery).ToList();
                            break;
                        case 2:
                            peripheries = peripheries.OrderByDescending(o => o.SumPeriphery).ToList();
                            break;
                    }

                    foreach (Periphery peripher in peripheries)
                    {
                        CurrentItem_listView.Items.Add(new CurrentItemControl(peripher, currentItem));
                    }
                    break;
            }
        }

        private void Find_textBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            CheckItem(CurrentItem);
        }

        private void CurrentItem_listView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (!isCheckys)
            {
                switch (CurrentItem)
                {
                    case 1:
                        CPU cpu = ((CurrentItemControl)CurrentItem_listView.SelectedItem).CPUID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(cpu, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 2:
                        MotherBoard motherBoard = ((CurrentItemControl)CurrentItem_listView.SelectedItem).MotherBoardID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(motherBoard, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 3:
                        Case @case = ((CurrentItemControl)CurrentItem_listView.SelectedItem).CaseID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(@case, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 4:
                        GPU gPU = ((CurrentItemControl)CurrentItem_listView.SelectedItem).GPUID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(gPU, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 5:
                        Cooler cooler = ((CurrentItemControl)CurrentItem_listView.SelectedItem).CoolerID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(cooler, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 6:
                        RAM RAM = ((CurrentItemControl)CurrentItem_listView.SelectedItem).RAMID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(RAM, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 7:
                        PowerSupply powerSupply = ((CurrentItemControl)CurrentItem_listView.SelectedItem).PowerSupplyID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(powerSupply, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 8:
                        Monitor monitor = ((CurrentItemControl)CurrentItem_listView.SelectedItem).MonitorID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(monitor, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 9:
                        Entity.Mouse mouse = ((CurrentItemControl)CurrentItem_listView.SelectedItem).MouseID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(mouse, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 10:
                        Entity.Keyboard keyboard = ((CurrentItemControl)CurrentItem_listView.SelectedItem).KeyboardID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(keyboard, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 11:
                        Computer computer = ((CurrentItemControl)CurrentItem_listView.SelectedItem).ComputerID;
                        ChooseWindow.currentWindow = this;
                        ChooseCurrentItemShow.ComputerReadyNot = computer;
                        new ChooseWindow(CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 12:
                        Periphery periphery = ((CurrentItemControl)CurrentItem_listView.SelectedItem).PeripheryID;
                        ChoosePeripheryWindow.currentWindow = this;
                        ChooseCurrentItemShow.PeripheryReadyNot = periphery;
                        new ChoosePeripheryWindow(CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                }
            }
            else
            {
                ProductInfoWindow.isCheckys = true;
                switch (CurrentItem)
                {
                    case 1:
                        CPU cpu = ((CurrentItemControl)CurrentItem_listView.SelectedItem).CPUID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(cpu, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 2:
                        MotherBoard motherBoard = ((CurrentItemControl)CurrentItem_listView.SelectedItem).MotherBoardID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(motherBoard, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 3:
                        Case @case = ((CurrentItemControl)CurrentItem_listView.SelectedItem).CaseID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(@case, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 4:
                        GPU gPU = ((CurrentItemControl)CurrentItem_listView.SelectedItem).GPUID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(gPU, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 5:
                        Cooler cooler = ((CurrentItemControl)CurrentItem_listView.SelectedItem).CoolerID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(cooler, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 6:
                        RAM RAM = ((CurrentItemControl)CurrentItem_listView.SelectedItem).RAMID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(RAM, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 7:
                        PowerSupply powerSupply = ((CurrentItemControl)CurrentItem_listView.SelectedItem).PowerSupplyID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(powerSupply, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 8:
                        Monitor monitor = ((CurrentItemControl)CurrentItem_listView.SelectedItem).MonitorID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(monitor, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 9:
                        Entity.Mouse mouse = ((CurrentItemControl)CurrentItem_listView.SelectedItem).MouseID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(mouse, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 10:
                        Entity.Keyboard keyboard = ((CurrentItemControl)CurrentItem_listView.SelectedItem).KeyboardID;
                        ProductInfoWindow.creatingForm = this;
                        new ProductInfoWindow(keyboard, CurrentItem).ShowDialog();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 11:
                        Computer computer = ((CurrentItemControl)CurrentItem_listView.SelectedItem).ComputerID;
                        ChooseWindow.currentWindow = this;
                        ChooseCurrentItemShow.ComputerReadyNot = computer;
                        new ChooseWindow(CurrentItem).ShowDialog();
                        ChooseCurrentItemShow.ComputerReadyNot = new Computer();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                    case 12:
                        Periphery periphery = ((CurrentItemControl)CurrentItem_listView.SelectedItem).PeripheryID;
                        ChoosePeripheryWindow.currentWindow = this;
                        ChooseCurrentItemShow.PeripheryReadyNot = periphery;
                        new ChoosePeripheryWindow(CurrentItem).ShowDialog();
                        ChooseCurrentItemShow.PeripheryReadyNot = new Periphery();
                        DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                        CheckItem(CurrentItem);
                        break;
                }
                ProductInfoWindow.isCheckys = false;
            }
        }

        private void BackWindow_button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Sorting_comboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CheckItem(CurrentItem);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Sorting_comboBox.Items.Add("Все");
            Sorting_comboBox.Items.Add("Упорядочить по цене ↑");
            Sorting_comboBox.Items.Add("Упорядочить по цене ↓");
            Sorting_comboBox.SelectedIndex = 0;
        }

        private void AddNewItem_button_Click(object sender, RoutedEventArgs e)
        {
            ProductInfoWindow.checkys = true;
            switch (CurrentItem)
            {
                
                case 1:
                    CPU cpu = new CPU();
                     new ProductInfoWindow(cpu, CurrentItem).ShowDialog();
                    DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                    CheckItem(CurrentItem);
                    break;
                case 2:
                    MotherBoard motherBoard = new MotherBoard();
                    ProductInfoWindow.creatingForm = this;
                    new ProductInfoWindow(motherBoard, CurrentItem).ShowDialog();
                    DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                    CheckItem(CurrentItem);
                    break;
                case 3:
                    Case @case = new Case();
                    ProductInfoWindow.creatingForm = this;
                    new ProductInfoWindow(@case, CurrentItem).ShowDialog();
                    DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                    CheckItem(CurrentItem);
                    break;
                case 4:
                    GPU gPU = new GPU();
                    ProductInfoWindow.creatingForm = this;
                    new ProductInfoWindow(gPU, CurrentItem).ShowDialog();
                    DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                    CheckItem(CurrentItem);
                    break;
                case 5:
                    Cooler cooler = new Cooler();
                    ProductInfoWindow.creatingForm = this;
                    new ProductInfoWindow(cooler, CurrentItem).ShowDialog();
                    DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                    CheckItem(CurrentItem);
                    break;
                case 6:
                    RAM RAM = new RAM();
                    ProductInfoWindow.creatingForm = this;
                    new ProductInfoWindow(RAM, CurrentItem).ShowDialog();
                    DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                    CheckItem(CurrentItem);
                    break;
                case 7:
                    PowerSupply powerSupply = new PowerSupply();
                    ProductInfoWindow.creatingForm = this;
                    new ProductInfoWindow(powerSupply, CurrentItem).ShowDialog();
                    DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                    CheckItem(CurrentItem);
                    break;
                case 8:
                    Monitor monitor = new Monitor();
                    ProductInfoWindow.creatingForm = this;
                    new ProductInfoWindow(monitor, CurrentItem).ShowDialog();
                    DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                    CheckItem(CurrentItem);
                    break;
                case 9:
                    Entity.Mouse mouse = new Entity.Mouse();
                    ProductInfoWindow.creatingForm = this;
                    new ProductInfoWindow(mouse, CurrentItem).ShowDialog();
                    DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                    CheckItem(CurrentItem);
                    break;
                case 10:
                    Entity.Keyboard keyboard = new Entity.Keyboard();
                    ProductInfoWindow.creatingForm = this;
                    new ProductInfoWindow(keyboard, CurrentItem).ShowDialog();
                    DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                    CheckItem(CurrentItem);
                    break;
                case 11:
                    new ChooseWindow().ShowDialog();
                    DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                    CheckItem(CurrentItem);
                    break;
                case 12:
                    new ChoosePeripheryWindow().ShowDialog();
                    DbContext = CreatePcForThePeopl_dbEntities.DBContext;
                    CheckItem(CurrentItem);
                    break;
            }
            ProductInfoWindow.checkys = false;
        }

        private void CurrentItem_listView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            if (CurrentItem_listView.SelectedIndex != -1)
            {
                AddToChoose_button.Visibility = Visibility.Visible;
            }
        }

        private void AddToChoose_button_Click(object sender, RoutedEventArgs e)
        {
           
            if (CurrentItem_listView.SelectedItem != null)
            {
                if (CurrentItem_listView.SelectedIndex != -1)
                {
                    if (!isCheckys)
                    { 
                        switch (CurrentItem)
                        {
                            case 1:

                                ComputerReadyNot.CPU = ((CurrentItemControl)CurrentItem_listView.SelectedItem).CPUID;
                                Close();
                                break;
                            case 2:
                                ComputerReadyNot.MotherBoard = ((CurrentItemControl)CurrentItem_listView.SelectedItem).MotherBoardID;
                                Close();
                                break;
                            case 3:
                                ComputerReadyNot.Case = ((CurrentItemControl)CurrentItem_listView.SelectedItem).CaseID;
                                Close();
                                break;
                            case 4:
                                ComputerReadyNot.GPU = ((CurrentItemControl)CurrentItem_listView.SelectedItem).GPUID;
                                Close();
                                break;
                            case 5:
                                ComputerReadyNot.Cooler = ((CurrentItemControl)CurrentItem_listView.SelectedItem).CoolerID;
                                Close();
                                break;
                            case 6:
                                ComputerReadyNot.RAM = ((CurrentItemControl)CurrentItem_listView.SelectedItem).RAMID;
                                Close();
                                break;
                            case 7:
                                ComputerReadyNot.PowerSupply = ((CurrentItemControl)CurrentItem_listView.SelectedItem).PowerSupplyID;
                                Close();
                                break;
                            case 8:
                                PeripheryReadyNot.Monitor = ((CurrentItemControl)CurrentItem_listView.SelectedItem).MonitorID;
                                Close();
                                break;
                            case 9:
                                PeripheryReadyNot.Mouse = ((CurrentItemControl)CurrentItem_listView.SelectedItem).MouseID;
                                Close();
                                break;
                            case 10:
                                PeripheryReadyNot.Keyboard = ((CurrentItemControl)CurrentItem_listView.SelectedItem).KeyboardID;
                                Close();
                                break;
                            case 11:
                                FinalStageOfOrder.OrderReadyNot.Computer = ((CurrentItemControl)CurrentItem_listView.SelectedItem).ComputerID;
                                Close();
                                break;
                            case 12:
                                FinalStageOfOrder.OrderReadyNot.Periphery = ((CurrentItemControl)CurrentItem_listView.SelectedItem).PeripheryID;
                                Close();
                                break;
                        }
                    }
                    else
                    {
                        switch (CurrentItem)
                        {
                            case 1:
                                if (MessageBox.Show("Вы точно хотите удалить этот товар", "Предупреждение", 
                                    MessageBoxButton.YesNo,MessageBoxImage.Hand)==MessageBoxResult.Yes)
                                {
                                    CPUID = ((CurrentItemControl)CurrentItem_listView.SelectedItem).CPUID;
                                    DbContext.CPU.Remove(CPUID);
                                    DbContext.SaveChanges();
                                    CheckItem(CurrentItem);
                                }
                                else
                                {
                                    break;
                                }
                                
                                break;
                            case 2:
                                if (MessageBox.Show("Вы точно хотите удалить этот товар", "Предупреждение",
                                    MessageBoxButton.YesNo, MessageBoxImage.Hand) == MessageBoxResult.Yes)
                                {
                                    MotherBoardID = ((CurrentItemControl)CurrentItem_listView.SelectedItem).MotherBoardID;
                                    DbContext.MotherBoard.Remove(MotherBoardID);
                                    DbContext.SaveChanges();
                                    CheckItem(CurrentItem);
                                }
                                else
                                {
                                    break;
                                }
                                break;
                            case 3:
                                if (MessageBox.Show("Вы точно хотите удалить этот товар", "Предупреждение",
                                    MessageBoxButton.YesNo, MessageBoxImage.Hand) == MessageBoxResult.Yes)
                                {
                                    CaseID = ((CurrentItemControl)CurrentItem_listView.SelectedItem).CaseID;
                                    DbContext.Case.Remove(CaseID);
                                    DbContext.SaveChanges();
                                    CheckItem(CurrentItem);
                                }
                                else
                                {
                                    break;
                                }
                                break;
                            case 4:
                                if (MessageBox.Show("Вы точно хотите удалить этот товар", "Предупреждение",
                                    MessageBoxButton.YesNo, MessageBoxImage.Hand) == MessageBoxResult.Yes)
                                {
                                    GPUID = ((CurrentItemControl)CurrentItem_listView.SelectedItem).GPUID;
                                    DbContext.GPU.Remove(GPUID);
                                    DbContext.SaveChanges();
                                    CheckItem(CurrentItem);
                                }
                                else
                                {
                                    break;
                                }
                                break;
                            case 5:
                                if (MessageBox.Show("Вы точно хотите удалить этот товар", "Предупреждение",
                                    MessageBoxButton.YesNo, MessageBoxImage.Hand) == MessageBoxResult.Yes)
                                {
                                    CoolerID = ((CurrentItemControl)CurrentItem_listView.SelectedItem).CoolerID;
                                    DbContext.Cooler.Remove(CoolerID);
                                    DbContext.SaveChanges();
                                    CheckItem(CurrentItem);
                                }
                                else
                                {
                                    break;
                                }
                                break;
                            case 6:
                                if (MessageBox.Show("Вы точно хотите удалить этот товар", "Предупреждение",
                                    MessageBoxButton.YesNo, MessageBoxImage.Hand) == MessageBoxResult.Yes)
                                {
                                    RAMID = ((CurrentItemControl)CurrentItem_listView.SelectedItem).RAMID;
                                    DbContext.RAM.Remove(RAMID);
                                    DbContext.SaveChanges();
                                    CheckItem(CurrentItem);
                                }
                                else
                                {
                                    break;
                                }
                                break;
                            case 7:
                                if (MessageBox.Show("Вы точно хотите удалить этот товар", "Предупреждение",
                                    MessageBoxButton.YesNo, MessageBoxImage.Hand) == MessageBoxResult.Yes)
                                {
                                    PowerSupplyID = ((CurrentItemControl)CurrentItem_listView.SelectedItem).PowerSupplyID;
                                    DbContext.PowerSupply.Remove(PowerSupplyID);
                                    DbContext.SaveChanges();
                                    CheckItem(CurrentItem);
                                }
                                else
                                {
                                    break;
                                }
                                break;
                            case 8:
                                if (MessageBox.Show("Вы точно хотите удалить этот товар", "Предупреждение",
                                    MessageBoxButton.YesNo, MessageBoxImage.Hand) == MessageBoxResult.Yes)
                                {
                                    MonitorID = ((CurrentItemControl)CurrentItem_listView.SelectedItem).MonitorID;
                                    DbContext.Monitor.Remove(MonitorID);
                                    DbContext.SaveChanges();
                                    CheckItem(CurrentItem);
                                }
                                else
                                {
                                    break;
                                }
                                break;
                            case 9:
                                if (MessageBox.Show("Вы точно хотите удалить этот товар", "Предупреждение",
                                    MessageBoxButton.YesNo, MessageBoxImage.Hand) == MessageBoxResult.Yes)
                                {
                                    MouseID = ((CurrentItemControl)CurrentItem_listView.SelectedItem).MouseID;
                                    DbContext.Mouse.Remove(MouseID);
                                    DbContext.SaveChanges();
                                    CheckItem(CurrentItem);
                                }
                                else
                                {
                                    break;
                                }
                                break;
                            case 10:
                                if (MessageBox.Show("Вы точно хотите удалить этот товар", "Предупреждение",
                                    MessageBoxButton.YesNo, MessageBoxImage.Hand) == MessageBoxResult.Yes)
                                {
                                    KeyboardID = ((CurrentItemControl)CurrentItem_listView.SelectedItem).KeyboardID;
                                    DbContext.Keyboard.Remove(KeyboardID);
                                    DbContext.SaveChanges();
                                    CheckItem(CurrentItem);
                                }
                                else
                                {
                                    break;
                                }
                                break;
                            case 11:
                                if (MessageBox.Show("Вы точно хотите удалить этот товар", "Предупреждение",
                                    MessageBoxButton.YesNo, MessageBoxImage.Hand) == MessageBoxResult.Yes)
                                {
                                    ComputerID = ((CurrentItemControl)CurrentItem_listView.SelectedItem).ComputerID;
                                    DbContext.Computer.Remove(ComputerID);
                                    DbContext.SaveChanges();
                                    CheckItem(CurrentItem);
                                }
                                else
                                {
                                    break;
                                }
                                break;
                            case 12:
                                if (MessageBox.Show("Вы точно хотите удалить этот товар", "Предупреждение",
                                    MessageBoxButton.YesNo, MessageBoxImage.Hand) == MessageBoxResult.Yes)
                                {
                                    PeripheryID = ((CurrentItemControl)CurrentItem_listView.SelectedItem).PeripheryID;
                                    DbContext.Periphery.Remove(PeripheryID);
                                    DbContext.SaveChanges();
                                    CheckItem(CurrentItem);
                                }
                                else
                                {
                                    break;
                                }
                                break;
                        }
                    }
                }
            }

        }
    }
}