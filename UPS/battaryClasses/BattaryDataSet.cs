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
            battaryDatas = ReadData();
        }
        public List<BattaryData> ReadData()
        {
            List<BattaryData>  battaryDatas = new List<BattaryData>();

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

                    if (readWorkbook != null)
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
                                            int id;
                                            if (int.TryParse(cell.Value.ToString(), out id))
                                                bat.ID = id;
                                            else
                                                bat.ID = 0;
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
                                            int cap;
                                            if (int.TryParse(cell.Value.ToString(), out cap)) //Проверка возможно ли пропарсить
                                                bat.Capacity = cap;
                                            else
                                                bat.Capacity = 0;
                                            break;
                                        case 8:
                                            int pow;
                                            if (int.TryParse(cell.Value.ToString(), out pow))
                                                bat.Power = pow;
                                            else
                                                bat.Power = 0;
                                            break;
                                        case 9:
                                            int volt;
                                            if (int.TryParse(cell.Value.ToString(), out volt))
                                                bat.Voltage = volt;
                                            else
                                                bat.Voltage = 0;
                                            break;
                                        case 10:
                                            double const2m;
                                            if (double.TryParse(cell.Value.ToString(), out const2m))
                                                bat.Const2m = const2m;
                                            else
                                                bat.Const2m = 0;
                                            break;
                                        case 11:
                                            double const4m;
                                            if (double.TryParse(cell.Value.ToString(), out const4m))
                                                bat.Const4m = const4m;
                                            else
                                                bat.Const4m = 0;
                                            break;
                                        case 12:
                                            double const5m;
                                            if (double.TryParse(cell.Value.ToString(), out const5m))
                                                bat.Const5m = const5m;
                                            else
                                                bat.Const5m = 0;
                                            break;
                                        case 13:
                                            double const6m;
                                            if (double.TryParse(cell.Value.ToString(), out const6m))
                                                bat.Const6m = const6m;
                                            else
                                                bat.Const6m = 0;
                                            break;
                                        case 14:
                                            double const8m;
                                            if (double.TryParse(cell.Value.ToString(), out const8m))
                                                bat.Const8m = const8m;
                                            else
                                                bat.Const8m = 0;
                                            break;
                                        case 15:
                                            double const10m;
                                            if (double.TryParse(cell.Value.ToString(), out const10m))
                                                bat.Const10m = const10m;
                                            else
                                                bat.Const10m = 0;
                                            break;
                                        case 16:
                                            double const15m;
                                            if (double.TryParse(cell.Value.ToString(), out const15m))
                                                bat.Const15m = const15m;
                                            else
                                                bat.Const15m = 0;
                                            break;
                                        case 17:
                                            double const20m;
                                            if (double.TryParse(cell.Value.ToString(), out const20m))
                                                bat.Const20m = const20m;
                                            else
                                                bat.Const20m = 0;
                                            break;
                                        case 18:
                                            double const30m;
                                            if (double.TryParse(cell.Value.ToString(), out const30m))
                                                bat.Const30m = const30m;
                                            else
                                                bat.Const30m = 0;
                                            break;
                                        case 19:
                                            double const45m;
                                            if (double.TryParse(cell.Value.ToString(), out const45m))
                                                bat.Const45m = const45m;
                                            else
                                                bat.Const45m = 0;
                                            break;
                                        case 20:
                                            double const60m;
                                            if (double.TryParse(cell.Value.ToString(), out const60m))
                                                bat.Const60m = const60m;
                                            else
                                                bat.Const60m = 0;
                                            break;
                                        case 21:
                                            double const90m;
                                            if (double.TryParse(cell.Value.ToString(), out const90m))
                                                bat.Const90m = const90m;
                                            else
                                                bat.Const90m = 0;
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
                                bat.FullDescrip = $"{bat.Mark}, {bat.Manufact}, {bat.Descrip}";
                                battaryDatas.Add(bat);
                            }
                            numOfRows++;
                        }
                        battaryDatas.RemoveAt(battaryDatas.Count - 1);
                    }
                }
                catch (Exception)
                {
                    ErrorWindow errorWindow = new ErrorWindow();
                    errorWindow.Show();
                    errorWindow.Topmost = true;
                    errorWindow.ErrorLable.Content = "Невозможно прочитать файл данных АКБ";
                }
            }
            return battaryDatas;
        }
        public object Clone() => this.MemberwiseClone();
    }
}
