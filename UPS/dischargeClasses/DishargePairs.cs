using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UPS.dischargeClasses
{
    public class DishargePairs
    {
        public int dischTime { get; set; } //Разрядное время
        public float dischPower { get; set; } //Мощность разрада на ячейку 2В, Вт/ячейку
    }
}
