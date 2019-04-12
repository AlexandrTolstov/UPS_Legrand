using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ClosedXML.Excel;

namespace UPS
{
    class BattaryData
    {
        readonly string name;
        readonly string artLeg;
        readonly string artOther;
        readonly string nameOtherBrand;
        readonly Dictionary<string, double> dischConst;

        //Наименование АКБ
        public string Name
        {
            get => name;
        }
        //Артикул АКБ Legrand
        public string ArtLeg
        {
            get => artLeg;
        }
        //Артикул АКБ Другого бренда
        public string ArtOther
        {
            get => artOther;
        }
        //Имя другого бренда
        public string NameOtherBrand
        {
            get => nameOtherBrand;
        }
        //Разрядная характеристика
        public Dictionary<string, double> DischConst
        {
            get => dischConst;
        }

        //Конструктор
        public BattaryData(
            string name, 
            string artLeg,
            string artOther,
            string nameOtherBrand,
            float const5m,
            float const10m,
            float const15m,
            float const20m,
            float const30m,
            float const45m,
            float const60m,
            float const2h,
            float const3h,
            float const5h,
            float const10h,
            float const20h)
        {
            this.name = name;
            this.artLeg = artLeg;
            this.artOther = artOther;
            this.nameOtherBrand = nameOtherBrand;
            dischConst = new Dictionary<string, double>();
            dischConst.Add("5 min", const5m);
            dischConst.Add("10 min", const10m);
            dischConst.Add("15 min", const15m);
            dischConst.Add("20 min", const20m);
            dischConst.Add("30 min", const30m);
            dischConst.Add("45 min", const45m);
            dischConst.Add("60 min", const60m);
            dischConst.Add("2 h", const2h);
            dischConst.Add("3 h", const3h);
            dischConst.Add("5 h", const5h);
            dischConst.Add("10 h", const10h);
            dischConst.Add("20 h", const20h);
        }

        public void WriteToFile()
        {
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

                //Запись файла
                string fileName = Name + ".xlsx";//Имя файла
                string fullPathName = workDirectory + folderData + "/" + fileName;//Полный путь с именем файла

                var worksheet = workbook.Worksheets.Add("Discharge constant");

                int i = 1;
                foreach(var pair in DischConst)
                {
                    worksheet.Column(i).Cell(1).Value = pair.Key;
                    worksheet.Column(i).Cell(2).Value = pair.Value;
                    i++;
                }
                i = 1;
                workbook.SaveAs(fullPathName);
            }
            
        }
    }
}
