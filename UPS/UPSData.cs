using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UPS
{
    class UPSData
    {
        public int ID { get; set; } //ID
        public string Name { get; set; } //Наименование ИБП
        public int KPD { get; set; } //КПД ИБП
        public int nBatLin { get; set; } //Количество батарей в линейке
        private int nElemBat { get; set; } //Количество элементов в батареи
        public UPSData(int ID, string Name, int KPD, int nBatLin)
        {
            this.ID = ID;
            this.Name = Name;
            this.KPD = KPD;
            this.nBatLin = nBatLin;
            this.nElemBat = 6;
        }
        public UPSData()
        {
            this.ID = 0;
            this.Name = "";
            this.KPD = 0;
            this.nBatLin = 0;
            this.nElemBat = 6;
        }
    }
}
