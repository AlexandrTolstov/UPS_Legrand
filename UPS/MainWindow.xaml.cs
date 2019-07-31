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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using ClosedXML.Excel;
using System.Collections.ObjectModel;
using UPS.dischargeClasses;

namespace UPS
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {      
        public MainWindow()
        {
            
            InitializeComponent();

            DataContext = new DataSource();

            //var dischList = new ObservableCollection<DishargePairs>();
            //dischList.Add(new DishargePairs() { dischPower = 0, dischTime = 0 });
            //this.dischGrid.ItemsSource = dischList;
            //this.dischTable.Children
        }
        internal sealed class DataSource
        {
            BattaryDataSet battaryDataSet = new BattaryDataSet();
            List<TimeSet> batTimes = new List<TimeSet>
            {
                new TimeSet(2),
                new TimeSet(4),
                new TimeSet(5),
                new TimeSet(6),
                new TimeSet(8),
                new TimeSet(10),
                new TimeSet(15),
                new TimeSet(20),
                new TimeSet(30),
                new TimeSet(45),
                new TimeSet(60),
                new TimeSet(90)
            };
            ListOfUPS listOfUPS = new ListOfUPS();
            public IEnumerable<BattaryData> BattaryList => battaryDataSet.ReadData(); //Считывает значения с BattaryDataSet
            public IEnumerable<BattaryData> BattaryList2 => battaryDataSet.ReadData(); //Считывает значения с BattaryDataSet
            public IEnumerable<UPSData> UPSList => listOfUPS.ReadData(); //Считывает значения с listOfUPS
            public IEnumerable<TimeSet> batTimeList => batTimes; //Время автономии
        }
        public void RaschetKolLineek()
        {
            int Pn; //Номинальная мощность, кВт
            if (!Int32.TryParse(PowerTextBox.Text, out Pn))
            {
                Pn = 0;
            }
            int KPDakb; //КПД АКБ
            int NbatLin; //Количество батарей в линейке

            if (KPDLable.Content != null && nBatLinLable.Content != null)
            {
                if (!Int32.TryParse(KPDLable.Content.ToString(), out KPDakb))
                {
                    KPDakb = 0;
                }
                if (!Int32.TryParse(nBatLinLable.Content.ToString(), out NbatLin))
                {
                    NbatLin = 0;
                }
            }
            else
            {
                KPDakb = 0;
                NbatLin = 0;
            }

            BattaryData selectedBattaryObj = (BattaryData)BatList.SelectedItem; //Выбранный АКБ объект
            TimeSet selectedTimeObj = (TimeSet)TimeList.SelectedItem; //Выбранное время объект

            int selectedTime; //Выбранное время
            double dischargeHaract; //Разрядная характеристика
            if (selectedBattaryObj != null && selectedTimeObj != null)
            {
                selectedTime = selectedTimeObj.time; //Выбранное время
                switch (selectedTime)
                {
                    case 2:
                        dischargeHaract = selectedBattaryObj.Const2m;
                        break;
                    case 4:
                        dischargeHaract = selectedBattaryObj.Const4m;
                        break;
                    case 5:
                        dischargeHaract = selectedBattaryObj.Const5m;
                        break;
                    case 6:
                        dischargeHaract = selectedBattaryObj.Const6m;
                        break;
                    case 8:
                        dischargeHaract = selectedBattaryObj.Const8m;
                        break;
                    case 10:
                        dischargeHaract = selectedBattaryObj.Const10m;
                        break;
                    case 15:
                        dischargeHaract = selectedBattaryObj.Const15m;
                        break;
                    case 20:
                        dischargeHaract = selectedBattaryObj.Const20m;
                        break;
                    case 30:
                        dischargeHaract = selectedBattaryObj.Const30m;
                        break;
                    case 45:
                        dischargeHaract = selectedBattaryObj.Const45m;
                        break;
                    case 60:
                        dischargeHaract = selectedBattaryObj.Const60m;
                        break;
                    case 90:
                        dischargeHaract = selectedBattaryObj.Const90m;
                        break;
                    default:
                        dischargeHaract = 1;
                        break;
                }
            }
            else
            {
                selectedTime = 0;
                dischargeHaract = 0;
            }
            
            float Pbat;
            float Pelem;
            double Nlin; //Количество линеек
            int Nlinround; //Количество линее округленное

            if (Pn > 0 && KPDakb > 0 && selectedBattaryObj != null && selectedTime > 0 && dischargeHaract > 0)
            {
                Pbat = (float)Pn * 100 / KPDakb; //Мощность батареи, кВт
                Pelem = Pbat * 1000 / (NbatLin * 6); //Мощность элемента в батарее, Вт
                Nlin = Pelem / dischargeHaract;
                Nlinround = (int)Math.Ceiling(Nlin);
                numOfLin.Content = Nlinround.ToString();
                numOfAKB.Content = (Nlinround * NbatLin).ToString();
            }
            else
            {
                numOfLin.Content = "Не определено";
                numOfAKB.Content = "Не определено";
            }
        }

        private void TimeList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RaschetKolLineek();
        }
        private void PowerTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            RaschetKolLineek();
        }
        private void UPSList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UPSData uPS = (UPSData)UPSList.SelectedItem;
            KPDLable.Content = uPS.KPD.ToString();
            nBatLinLable.Content = uPS.nBatLin.ToString();
            RaschetKolLineek();
        }
        private void BatList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BattaryData batData = (BattaryData)BatList.SelectedItem;
            if (batteryGrid.Items.Count > 0)
                batteryGrid.Items.RemoveAt(0);
            batteryGrid.Items.Add(batData);
            RaschetKolLineek();
        }

        private void AddColumn_Click(object sender, RoutedEventArgs e)
        {
            //var bordBrush1 = new Border { BorderThickness = new Thickness(1.0), BorderBrush = Brushes.Black };
            var textBox1 = new TextBox { Width = 80, BorderThickness = new Thickness(1.0), BorderBrush = Brushes.Black };
            var textBox2 = new TextBox { Width = 80, BorderThickness = new Thickness(1.0), BorderBrush = Brushes.Black };

            dischTable.ColumnDefinitions.Add(new ColumnDefinition());
            dischTable.Children.Add(textBox1);
            dischTable.Children.Add(textBox2);

            Grid.SetColumn(textBox1, dischTable.ColumnDefinitions.Count - 1);
            Grid.SetColumn(textBox2, dischTable.ColumnDefinitions.Count - 1);
            Grid.SetRow(textBox1, 0);
            Grid.SetRow(textBox2, 1);
        }

        private void RemuveColumn_Click(object sender, RoutedEventArgs e)
        {
            if (dischTable.ColumnDefinitions.Count > 2)
            {
                dischTable.ColumnDefinitions.RemoveAt(dischTable.ColumnDefinitions.Count - 1);
                dischTable.Children.RemoveAt(dischTable.Children.Count - 1);
                dischTable.Children.RemoveAt(dischTable.Children.Count - 1);
            }
        }

        private void GetValue_Click(object sender, RoutedEventArgs e)
        {
            string txt = "";
            DischTableClass.AddPairs(dischTable);
            foreach (var item in DischTableClass.dishargePairs)
            {
                if(item.dischPower == 0 || item.dischTime == 0)
                {
                    MessageBox.Show("Вы не заполнили все ячеки");
                    break;
                }
            }
            foreach (var item in DischTableClass.dishargePairs)
            {
                txt += item.dischTime.ToString();
                txt += "-";
                txt += item.dischPower.ToString();
                txt += "\n";
            }
            InfoLable.Content = txt;
        }    
        public string GetValueGrid(int row, int col)//Функция получения значений из Grid по номеру ряда и столбца
        {
            string txt = "";
            if (row >= 0 && col >= 0)
            {
                var ch = dischTable.Children.Cast<UIElement>().First(b => Grid.GetRow(b) == row && Grid.GetColumn(b) == col + 1);
                txt = (((TextBox)ch).Text).ToString();
            }
            else txt = "не верный диапозон";
            return txt;
        }

        private void StackPanel_MouseEnter(object sender, MouseEventArgs e)
        {
            (sender as StackPanel).Background = Brushes.Yellow;
        }

        private void StackPanel_MouseLeave(object sender, MouseEventArgs e)
        {
            (sender as StackPanel).Background = Brushes.LightSeaGreen;
        }
    }
}
