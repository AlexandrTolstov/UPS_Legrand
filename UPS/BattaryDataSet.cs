using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace UPS
{
    public class BattaryDataSet:ICloneable
    {
        public List<BattaryData> battaryDatas { get; set; } //Перечень АКБ
        public BattaryDataSet()
        {
            battaryDatas = new List<BattaryData>();

            using (var workbook = new XLWorkbook())
            {
                string workDirectory = Environment.CurrentDirectory;//Читает путь с файлом exe
                string folderData = "/data";//Имя папки

                //В случае если папка data отсутствует создаем ее
                DirectoryInfo dirInfo = new DirectoryInfo(workDirectory + folderData);
                if (!dirInfo.Exists)
                {
                    dirInfo.Create();
                }

                string Name = "BatteryData";
                //Запись файла
                string fileName = Name + ".xlsx";//Имя файла
                string fullPathName = workDirectory + folderData + "/" + fileName;//Полный путь с именем файла

                try
                {
                    var readWorkbook = new XLWorkbook(fullPathName);

                    if(readWorkbook != null)
                    {
                        var worksheet = readWorkbook.Worksheet(1);

                        int numOfRows = 1;
                        foreach (var row in worksheet.Rows())
                        {
                            int numOfCell = 1;
                            if (numOfRows >= 3)
                            {
                                BattaryData bat = new BattaryData();
                                foreach (var cell in row.Cells())
                                {
                                    switch (numOfCell)
                                    {
                                        case 1:
                                            bat.ID = int.Parse(cell.Value.ToString());
                                            break;
                                        case 2:
                                            bat.LegArt = cell.Value.ToString();
                                            break;
                                        case 3:
                                            bat.OtherArt = cell.Value.ToString();
                                            break;
                                        case 4:
                                            bat.Mark = cell.Value.ToString();
                                            break;
                                        case 5:
                                            bat.Manufact = cell.Value.ToString();
                                            break;
                                        case 6:
                                            bat.Descrip = cell.Value.ToString();
                                            break;
                                        case 7:
                                            bat.Capacity = int.Parse(cell.Value.ToString());
                                            break;
                                        case 8:
                                            bat.Power = int.Parse(cell.Value.ToString());
                                            break;
                                        case 9:
                                            bat.Voltage = int.Parse(cell.Value.ToString());
                                            break;
                                        case 10:
                                            bat.Const2m = double.Parse(cell.Value.ToString());
                                            break;
                                        case 11:
                                            bat.Const4m = double.Parse(cell.Value.ToString());
                                            break;
                                        case 12:
                                            bat.Const5m = double.Parse(cell.Value.ToString());
                                            break;
                                        case 13:
                                            bat.Const6m = double.Parse(cell.Value.ToString());
                                            break;
                                        case 14:
                                            bat.Const8m = double.Parse(cell.Value.ToString());
                                            break;
                                        case 15:
                                            bat.Const10m = double.Parse(cell.Value.ToString());
                                            break;
                                        case 16:
                                            bat.Const15m = double.Parse(cell.Value.ToString());
                                            break;
                                        case 17:
                                            bat.Const20m = double.Parse(cell.Value.ToString());
                                            break;
                                        case 18:
                                            bat.Const30m = double.Parse(cell.Value.ToString());
                                            break;
                                        case 19:
                                            bat.Const45m = double.Parse(cell.Value.ToString());
                                            break;
                                        case 20:
                                            bat.Const60m = double.Parse(cell.Value.ToString());
                                            break;
                                        case 21:
                                            bat.Const90m = double.Parse(cell.Value.ToString());
                                            break;
                                        case 22:
                                            bat.LegArt_Euro = cell.Value.ToString();
                                            break;
                                        case 23:
                                            bat.OtherArt_Euro = cell.Value.ToString();
                                            break;
                                        case 24:
                                            bat.Descrip_Euro = cell.Value.ToString();
                                            break;
                                        default:
                                            break;
                                    }
                                    numOfCell++;
                                }
                                battaryDatas.Add(bat);
                            }
                            numOfRows++;
                        }
                        battaryDatas.RemoveAt(battaryDatas.Count - 1);
                    }
                }
                catch(Exception)
                {
                    ErrorWindow errorWindow = new ErrorWindow();
                    errorWindow.Show();
                    errorWindow.Topmost = true;
                    errorWindow.ErrorLable.Content = "Невозможно прочитать файл данных АКБ";
                }
            }
        }

        //public object Clone()
        //{
        //    BattaryDataSet newBatDatSet = (BattaryDataSet)this.MemberwiseClone();
        //    List<BattaryData> tempBatDataSet = new List<BattaryData>();

        //    foreach (var it in this.battaryDatas)
        //    {
        //        BattaryData temBatData = (BattaryData)it.Clone();
        //        tempBatDataSet.Add(temBatData);
        //    }
        //    newBatDatSet.battaryDatas = tempBatDataSet;
        //    return newBatDatSet;
        //}

        public object Clone() => this.MemberwiseClone();
    }
}
