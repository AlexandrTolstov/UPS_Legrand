using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ClosedXML.Excel;

namespace UPS
{
    public class BattaryData
    {
        public int ID { get; set; } //ID АКБ
        public string LegArt { get; set; } //Артикул Legrand
        public string OtherArt { get; set; } //Артикул другого производителя
        public string Mark { get; set; } //Марка
        public string Manufact { get; set; } //Производитель
        public string Descrip { get; set; } //Описание
        public int Capacity { get; set; } //Емкость
        public int Power { get; set; } //Мощность
        public int Voltage { get; set; } //Напряжение

        //Разрядная характеристика
        public double Const2m { get; set; } //2 min
        public double Const4m { get; set; } //4 min
        public double Const5m { get; set; } //5 min
        public double Const6m { get; set; } //6 min
        public double Const8m { get; set; } //8 min
        public double Const10m { get; set; } //10 min
        public double Const15m { get; set; } //15 min
        public double Const20m { get; set; } //20 min
        public double Const30m { get; set; } //30 min
        public double Const45m { get; set; } //45 min
        public double Const60m { get; set; } //60 min
        public double Const90m { get; set; } //90 min

        //Конструктор
        public BattaryData(
            int ID, 
            string LegArt,
            string OtherArt,
            string Mark,
            string Manufact,
            string Descrip,
            int Capacity,
            int Power,
            int Voltage,
            double Const2m,
            double Const4m,
            double Const5m,
            double Const6m,
            double Const8m,
            double Const10m,
            double Const15m,
            double Const20m,
            double Const30m,
            double Const45m,
            double Const60m,
            double Const90m)
        {
            this.ID = ID;
            this.LegArt = LegArt;
            this.OtherArt = OtherArt;
            this.Mark = Mark;
            this.Manufact = Manufact;
            this.Descrip = Descrip;
            this.Capacity = Capacity;
            this.Power = Power;
            this.Voltage = Voltage;
            this.Const2m = Const2m;
            this.Const4m = Const4m;
            this.Const5m = Const5m;
            this.Const6m = Const6m;
            this.Const8m = Const8m;
            this.Const10m = Const10m;
            this.Const15m = Const15m;
            this.Const20m = Const20m;
            this.Const30m = Const30m;
            this.Const45m = Const45m;
            this.Const60m = Const60m;
            this.Const90m = Const90m;
        }

        public BattaryData()
        {
            this.ID = 0;
            this.LegArt = "";
            this.OtherArt = "";
            this.Mark = "";
            this.Manufact = "";
            this.Descrip = "";
            this.Capacity = 0;
            this.Power = 0;
            this.Voltage = 0;
            this.Const2m = 0;
            this.Const4m = 0;
            this.Const5m = 0;
            this.Const6m = 0;
            this.Const8m = 0;
            this.Const10m = 0;
            this.Const15m = 0;
            this.Const20m = 0;
            this.Const30m = 0;
            this.Const45m = 0;
            this.Const60m = 0;
            this.Const90m = 0;
        }
        //public void WriteToFile()
        //{
        //    using (var workbook = new XLWorkbook())
        //    {
        //        string workDirectory = Environment.CurrentDirectory;//Читает путь с файлом exe
        //        string folderData = "/data";//Имя папки

        //        //В случае если папка data отсутствует создаем ее
        //        DirectoryInfo dirInfo = new DirectoryInfo(workDirectory + folderData);
        //        if (!dirInfo.Exists)
        //        {
        //            dirInfo.Create();
        //        }

        //        //Запись файла
        //        string fileName = Name + ".xlsx";//Имя файла
        //        string fullPathName = workDirectory + folderData + "/" + fileName;//Полный путь с именем файла

        //        var worksheet = workbook.Worksheets.Add("Discharge constant");

        //        int i = 1;
        //        foreach(var pair in DischConst)
        //        {
        //            worksheet.Column(i).Cell(1).Value = pair.Key;
        //            worksheet.Column(i).Cell(2).Value = pair.Value;
        //            i++;
        //        }
        //        i = 1;
        //        workbook.SaveAs(fullPathName);
        //    }
            
        //}
    }

}
