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
            //using(var workbook = new XLWorkbook())
            //{
            //    string workDirectory = Environment.CurrentDirectory;//Читает путь с файлом exe
            //    string folderData = "/data";//Имя папки

            //    //В случае если папка data отсутствует создаем ее
            //    DirectoryInfo dirInfo = new DirectoryInfo(workDirectory + folderData);
            //    if (!dirInfo.Exists)
            //    {
            //        dirInfo.Create();
            //    }

            //    //Запись файла
            //    string fileName = "HelloWorld.xlsx";//Имя файла
            //    string fullPathName = workDirectory + folderData + "/" + fileName;//Полный путь с именем файла

            //    var worksheet = workbook.Worksheets.Add("Sample Sheet");
            //    worksheet.Cell("A1").Value = "Hello Worldssss!";
            //    worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
            //    worksheet.Cell("A20").Value = "Эх ма";


            //    workbook.SaveAs(fullPathName);

            //    //Чтение файла
            //    var readWorkbook = new XLWorkbook(fullPathName);
            //    var readWorksheet = readWorkbook.Worksheet(1);
            //    //var rows = worksheet.RangeUsed().RowsUsed();

            //    //Вывод 

            //    //Вариант 1
            //    //Label1.Content = readWorksheet.Cell("A20").Value; 

            //    //Вариант 2
            //    //Label1.Content = readWorksheet.Column(1).Cell(20).Value;        
            //}


            //BattaryDataSet battaryDataSet = new BattaryDataSet();

            DataContext = new DataSource();
        }
        internal sealed class DataSource
        {
            BattaryDataSet battaryDataSet = new BattaryDataSet();
            ListOfUPS listOfUPS = new ListOfUPS();
            public IEnumerable<BattaryData> BattaryList => battaryDataSet.battaryDatas; //Считывает значения с BattaryDataSet
            public IEnumerable<UPSData> UPSList => listOfUPS.UPSDatas; //Считывает значения с listOfUPS
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
            TypeOfObj.Content = batData.Manufact.ToString();
        }
    }
}
