using System.Windows;

namespace PcSborka.Views.Windows
{
    /// <summary>
    /// Логика взаимодействия для OrderCreate.xaml
    /// </summary>
    public partial class OrderCreate : Window
    {
        public OrderCreate()
        {
            InitializeComponent();
        }

        private void Back_button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ChooseComputer_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window chooseComputerWindow = new ChooseWindow();
            chooseComputerWindow.ShowDialog();
            Show();
            if (ChoosePeripheryWindow.countPeriphery > 0)
            {
                if (ChooseCurrentItemShow.ComputerReadyNot.PowerSupply != null)
                {
                    ComponentsInfo_textblock.Text = $"Сумма стоимости комплектующих:{ChooseCurrentItemShow.ComputerReadyNot.SumComponents} руб.\nВы успешно cобрали комплекующие, \n теперь вы можете сделать заказ";
                    ChooseComputer_button.IsEnabled = false;

                    Periphery_textBlock.Text = "Или выбрать периферию";

                    ChooseFinishedComputer_button.Visibility = Visibility.Hidden;
                    ChooseFinishedPeriphery_button.Visibility = Visibility.Hidden;
                    FinishedInfo_textBlock.Visibility = Visibility.Hidden;

                    CreateOrder_button.Visibility = Visibility.Visible;
                    ChooseComputer_button.IsEnabled = false;
                }
                else
                {
                    if (ChooseCurrentItemShow.ComputerReadyNot.isBuildComputer == true)
                    {
                        ComponentsInfo_textblock.Text = $"Сумма примерной стоимости комплектующих:{ChooseCurrentItemShow.ComputerReadyNot.SumComponents} руб.\nВы успешно выбрали комплекующие,\nих вам дособерут сотрудники ,\n теперь вы можете сделать заказ";

                        ChooseComputer_button.IsEnabled = false;

                        Periphery_textBlock.Text = "Или выбрать периферию";

                        ChooseFinishedComputer_button.Visibility = Visibility.Hidden;
                        ChooseFinishedPeriphery_button.Visibility = Visibility.Hidden;
                        FinishedInfo_textBlock.Visibility = Visibility.Hidden;

                        CreateOrder_button.Visibility = Visibility.Visible;
                        ChooseComputer_button.IsEnabled = false;
                    }
                    else
                    {
                        ChooseCurrentItemShow.ComputerReadyNot = new Entity.Computer();
                    }
                }
            }
            else
            {
                if (ChooseCurrentItemShow.ComputerReadyNot.PowerSupply != null)
                {
                    ComponentsInfo_textblock.Text = $"Сумма стоимости комплектующих:{ChooseCurrentItemShow.ComputerReadyNot.SumComponents} руб.\nВы успешно cобрали комплекующие, \n теперь вы можете сделать заказ";
                    ChooseComputer_button.IsEnabled = false;

                    ChooseFinishedComputer_button.Visibility = Visibility.Hidden;
                    ChooseFinishedPeriphery_button.Visibility = Visibility.Hidden;
                    FinishedInfo_textBlock.Visibility = Visibility.Hidden;

                    CreateOrder_button.Visibility = Visibility.Visible;
                    ChooseComputer_button.IsEnabled = false;
                }
                else
                {
                    if (ChooseCurrentItemShow.ComputerReadyNot.isBuildComputer == true)
                    {
                        ComponentsInfo_textblock.Text = $"Сумма примерной стоимости комплектующих:{ChooseCurrentItemShow.ComputerReadyNot.SumComponents} руб.\nВы успешно выбрали комплекующие,\nих вам дособерут сотрудники ,\n теперь вы можете сделать заказ";

                        ChooseComputer_button.IsEnabled = false;


                        ChooseFinishedComputer_button.Visibility = Visibility.Hidden;
                        ChooseFinishedPeriphery_button.Visibility = Visibility.Hidden;
                        FinishedInfo_textBlock.Visibility = Visibility.Hidden;

                        CreateOrder_button.Visibility = Visibility.Visible;
                        ChooseComputer_button.IsEnabled = false;
                    }
                    else
                    {
                        ChooseCurrentItemShow.ComputerReadyNot = new Entity.Computer();
                    }
                }
            }

        }
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ChooseCurrentItemShow.ComputerReadyNot = new Entity.Computer();
            ChooseCurrentItemShow.PeripheryReadyNot = new Entity.Periphery();
            FinalStageOfOrder.OrderReadyNot = new Entity.Order();
        }

        private void ChoosePeriphery_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window choosePeriphery = new ChoosePeripheryWindow();
            choosePeriphery.ShowDialog();
            Show();
            if (ChooseCurrentItemShow.ComputerReadyNot.CPU == null)
            {
                if (ChoosePeripheryWindow.countPeriphery >0)
                {
                    Periphery_textBlock.Text = $"Сумма стоимости периферий:{ChooseCurrentItemShow.PeripheryReadyNot.SumPeriphery} руб.\nВы успешно cобрали периферию";
                    ChoosePeriphery_button.IsEnabled = false;

                    ComponentsInfo_textblock.Text = "Теперь вы должны собрать комплектующие";

                    ChooseFinishedComputer_button.Visibility = Visibility.Hidden;
                    ChooseFinishedPeriphery_button.Visibility = Visibility.Hidden;
                    FinishedInfo_textBlock.Visibility = Visibility.Hidden;


                    ChoosePeripheryWindow.countPeriphery = 0;
                }
                else
                {
                    if (ChooseCurrentItemShow.PeripheryReadyNot.isBuildPeriphery == true)
                    {
                        Periphery_textBlock.Text = $"Сумма стоимости периферий:{ChooseCurrentItemShow.PeripheryReadyNot.SumPeriphery} руб.,\nих вам дособерут сотрудники";
                        ChoosePeriphery_button.IsEnabled = false;

                        ComponentsInfo_textblock.Text = "Теперь вы должны собрать комплектующие";

                        ChooseFinishedComputer_button.Visibility = Visibility.Hidden;
                        ChooseFinishedPeriphery_button.Visibility = Visibility.Hidden;
                        FinishedInfo_textBlock.Visibility = Visibility.Hidden;

                        ChooseComputer_button.IsEnabled = false;

                        ChoosePeripheryWindow.countPeriphery = 0;
                    }
                    else
                    {
                        ChooseCurrentItemShow.PeripheryReadyNot = new Entity.Periphery();

                        ChoosePeripheryWindow.countPeriphery = 0;
                    }
                }
            }
            else
            {
                if (ChoosePeripheryWindow.countPeriphery > 0)
                {
                    Periphery_textBlock.Text = $"Сумма стоимости периферий:{ChooseCurrentItemShow.PeripheryReadyNot.SumPeriphery} руб.\nВы успешно cобрали периферию";
                    ChoosePeriphery_button.IsEnabled = false;

                    ChoosePeripheryWindow.countPeriphery = 0;
                }
                else
                {
                    if (ChooseCurrentItemShow.PeripheryReadyNot.isBuildPeriphery == true)
                    {
                        Periphery_textBlock.Text = $"Сумма стоимости периферий:{ChooseCurrentItemShow.PeripheryReadyNot.SumPeriphery} руб.,\nих вам дособерут сотрудники";
                        ChoosePeriphery_button.IsEnabled = false;

                        ChoosePeripheryWindow.countPeriphery = 0;
                    }
                    else
                    {
                        ChooseCurrentItemShow.PeripheryReadyNot = new Entity.Periphery();

                        ChoosePeripheryWindow.countPeriphery = 0;
                    }
                }
            }
        }

        private void ChooseFinishedComputer_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window currentItemShowWindow = new ChooseCurrentItemShow(11);
            currentItemShowWindow.ShowDialog();
            Show();
            if (FinalStageOfOrder.OrderReadyNot.Periphery == null)
            {
                if (FinalStageOfOrder.OrderReadyNot.Computer != null)
                {
                    ComponentsInfo_textblock.Visibility = Visibility.Hidden;

                    Periphery_textBlock.Visibility = Visibility.Hidden;

                    ChooseComputer_button.Visibility = Visibility.Hidden;

                    ChoosePeriphery_button.Visibility = Visibility.Hidden;

                    CreateOrder_button.Visibility = Visibility.Visible;

                    ChooseFinishedComputer_button.IsEnabled = false;

                    FinishedInfo_textBlock.Text = $"Вы успешно выбрали готовую сборку компьютера\nСумма комплектующих:" +
                        $"{FinalStageOfOrder.OrderReadyNot.Computer.SumComponents} руб.\nТеперь вы можете сделать заказ\nИли выбрать готовую сборку периферий";
                }
            }
            else
            {
                if (FinalStageOfOrder.OrderReadyNot.Computer != null)
                {

                    CreateOrder_button.Visibility = Visibility.Visible;

                    ChooseFinishedComputer_button.IsEnabled = false;

                    FinishedInfo_textBlock.Text = $"Вы успешно выбрали готовую сборку компьютера и периферий\nСумма комплектующих компьютера:" +
                        $"{FinalStageOfOrder.OrderReadyNot.Computer.SumComponents} руб.\nСумма периферий: {FinalStageOfOrder.OrderReadyNot.Periphery.SumPeriphery} руб.\nТеперь вы можете сделать заказ";
                }
            }
        }

        private void CreateOrder_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window nextWindow = new FinalStageOfOrder();
            nextWindow.ShowDialog();
            Show();
            if (FinalStageOfOrder.isChecked != false)
            {
                Close();
            }
        }

        private void ChooseFinishedPeriphery_button_Click(object sender, RoutedEventArgs e)
        {
            Hide();
            Window currentItemShowWindow = new ChooseCurrentItemShow(12);
            currentItemShowWindow.ShowDialog();
            Show();
            if(FinalStageOfOrder.OrderReadyNot.Computer == null)
            {
                if(FinalStageOfOrder.OrderReadyNot.Periphery !=null)
                {
                    ComponentsInfo_textblock.Visibility = Visibility.Hidden;

                    Periphery_textBlock.Visibility = Visibility.Hidden;

                    ChooseComputer_button.Visibility = Visibility.Hidden;

                    ChoosePeriphery_button.Visibility = Visibility.Hidden;

                    ChooseFinishedPeriphery_button.IsEnabled = false;

                    FinishedInfo_textBlock.Text = $"Вы успешно выбрали готовую сборку периферий\nСумма периферий: {FinalStageOfOrder.OrderReadyNot.Periphery.SumPeriphery} руб.\nТеперь вы должны выбрать готовую сборку компьютера";
                }
            }
            else
            {
                FinishedInfo_textBlock.Text = $"Вы успешно выбрали готовую сборку компьютера и периферий\nСумма комплектующих компьютера:" +
                     $"{FinalStageOfOrder.OrderReadyNot.Computer.SumComponents} руб.\nСумма периферий: {FinalStageOfOrder.OrderReadyNot.Periphery.SumPeriphery} руб.\nТеперь вы можете сделать заказ";
                ChooseFinishedPeriphery_button.IsEnabled = false;
            }
        }
    }
}
