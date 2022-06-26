using PcSborka.Entity;
using System.Windows;
using System.Windows.Controls;

namespace PcSborka.Views.UserControls
{
    /// <summary>
    /// Логика взаимодействия для CurrentItemControl.xaml
    /// </summary>
    public partial class CurrentItemControl : UserControl
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

        #region Constructs
        public CurrentItemControl(CPU cpu, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            CPUID = cpu;
            CurrentItem = currentItem;
        }

        public CurrentItemControl(MotherBoard motherBoard, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            MotherBoardID = motherBoard;
            CurrentItem = currentItem;
        }

        public CurrentItemControl(Case @case, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            CaseID = @case;
            CurrentItem = currentItem;
        }

        public CurrentItemControl(GPU gPU, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            GPUID = gPU;
            CurrentItem = currentItem;
        }

        public CurrentItemControl(Cooler cooler, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            CoolerID = cooler;
            CurrentItem = currentItem;
        }

        public CurrentItemControl(RAM rAM, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            RAMID = rAM;
            CurrentItem = currentItem;
        }

        public CurrentItemControl(PowerSupply powerSupply, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            PowerSupplyID = powerSupply;
            CurrentItem = currentItem;
        }

        public CurrentItemControl(Monitor monitor, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            MonitorID = monitor;
            CurrentItem = currentItem;
        }

        public CurrentItemControl(Entity.Mouse mouse, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            MouseID = mouse;
            CurrentItem = currentItem;
        }

        public CurrentItemControl(Entity.Keyboard keyboard, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            KeyboardID = keyboard;
            CurrentItem = currentItem;
        }

        public CurrentItemControl(Computer computer, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            ComputerID = computer;
            CurrentItem = currentItem;
        }

        public CurrentItemControl(Periphery periphery, int currentItem)
        {
            InitializeComponent();
            DbContext = CreatePcForThePeopl_dbEntities.DBContext;
            PeripheryID = periphery;
            CurrentItem = currentItem;
        }
        #endregion
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            CheckCurrentItem(CurrentItem);
        }

        private void CheckCurrentItem(int currentItem)
        {
            switch (currentItem)
            {
                //pivo
                case 1:
                    Title_textBlock.Text = CPUID.Name;
                    Cost_textBlock.Text = $"Цена: {CPUID.Cost} руб.";
                    break;
                case 2:
                    Title_textBlock.Text = MotherBoardID.Name;
                    Cost_textBlock.Text = $"Цена: {MotherBoardID.Cost} руб.";
                    break;
                case 3:
                    Title_textBlock.Text = CaseID.Name;
                    Cost_textBlock.Text = $"Цена: {CaseID.Cost} руб.";
                    break;
                case 4:
                    Title_textBlock.Text = GPUID.Name;
                    Cost_textBlock.Text = $"Цена: {GPUID.Cost} руб.";
                    break;
                case 5:
                    Title_textBlock.Text =CoolerID.Name;
                    Cost_textBlock.Text = $"Цена: {CoolerID.Cost} руб.";
                    break;
                case 6:
                    Title_textBlock.Text = RAMID.Name;
                    Cost_textBlock.Text = $"Цена: {RAMID.Cost} руб.";
                    break;
                case 7:
                    Title_textBlock.Text = PowerSupplyID.Name;
                    Cost_textBlock.Text = $"Цена: {PowerSupplyID.Cost} руб.";
                    break;
                case 8:
                    Title_textBlock.Text = MonitorID.Name;
                    Cost_textBlock.Text = $"Цена: {MonitorID.Cost} руб.";
                    break;
                case 9:
                    Title_textBlock.Text = MouseID.Name;
                    Cost_textBlock.Text = $"Цена: {MouseID.Cost} руб.";
                    break;
                case 10:
                    Title_textBlock.Text = KeyboardID.Name;
                    Cost_textBlock.Text = $"Цена: {KeyboardID.Cost} руб.";
                    break;
                case 11:
                    Title_textBlock.Text = $"Сборка компьютера номер #{ComputerID.ID}";
                    Cost_textBlock.Text = $"Цена: {ComputerID.SumComponents} руб.";
                    break;
                case 12:
                    Title_textBlock.Text = $"Сборка периферий номер #{PeripheryID.ID}";
                    Cost_textBlock.Text = $"Цена: {PeripheryID.SumPeriphery} руб.";
                    break;
            }
        }
    }
}
