using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UPS
{
    class ListOfUPS
    {
        public List<UPSData> UPSDatas { get; set; } //Перечень ИБП
        public ListOfUPS()
        {
            UPSDatas = new List<UPSData>();

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

                string Name = "UPSData";
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
                            if (numOfRows >= 2)
                            {
                                int num=0;
                                UPSData ups = new UPSData();
                                foreach (var cell in row.Cells())
                                {
                                    switch (numOfCell)
                                    {
                                        case 1:
                                            ups.ID = int.TryParse(cell.Value.ToString(), out num) ? num : 0;
                                            break;
                                        case 2:
                                            ups.Name = cell.Value.ToString();
                                            break;
                                        case 3:
                                            ups.KPD = int.TryParse(cell.Value.ToString(), out num) ? num : 0;
                                            break;
                                        case 4:
                                            ups.nBatLin = int.TryParse(cell.Value.ToString(), out num) ? num : 0;
                                            break;
                                        default:
                                            break;
                                    }
                                    numOfCell++;
                                }
                                UPSDatas.Add(ups);
                            }
                            numOfRows++;
                        }
                        UPSDatas.RemoveAt(UPSDatas.Count - 1);
                    }
                }
                catch (Exception ex)
                {
                    ErrorWindow errorWindow = new ErrorWindow();
                    errorWindow.Show();
                    errorWindow.Topmost = true;
                    errorWindow.ErrorLable.Content = ex.ToString();
                }
            }
        }
    }
}
