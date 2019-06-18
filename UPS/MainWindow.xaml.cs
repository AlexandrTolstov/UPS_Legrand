﻿using System;
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
        }
        internal sealed class DataSource
        {
            BattaryDataSet battaryDataSet = new BattaryDataSet();
            BattaryDataSet battaryDataSet2 = new BattaryDataSet();
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
            public IEnumerable<BattaryData> BattaryList => battaryDataSet.battaryDatas; //Считывает значения с BattaryDataSet
            public IEnumerable<BattaryData> BattaryList2 => battaryDataSet2.battaryDatas; //Считывает значения с BattaryDataSet
            public IEnumerable<UPSData> UPSList => listOfUPS.UPSDatas; //Считывает значения с listOfUPS
            public IEnumerable<TimeSet> batTimeList => batTimes; //Время автономии
        }
        private void UPSList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UPSData uPS = (UPSData)UPSList.SelectedItem;
            KPDLable.Content = uPS.KPD.ToString();
            nBatLinLable.Content = uPS.nBatLin.ToString();
        }

        private void BatList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BattaryData batData = (BattaryData)BatList.SelectedItem;
            if (batteryGrid.Items.Count > 0)
                batteryGrid.Items.RemoveAt(0);
            batteryGrid.Items.Add(batData);
        }

        private void TimeList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}
