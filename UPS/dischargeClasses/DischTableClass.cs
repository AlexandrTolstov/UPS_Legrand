using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace UPS.dischargeClasses
{
    public static class DischTableClass
    {
        public static List<DishargePairs> dishargePairs = new List<DishargePairs>();
        //public static void AddPair(DishargePairs pairs)
        //{
        //    dishargePairs.Add(pairs);
        //}
        public static void AddPairs(Grid table)
        {
            string txt1 = "";
            string txt2 = "";
            int time;
            float power;
            if (dishargePairs.Count != 0)
                dishargePairs.Clear();
            for (int i = 0; i < table.ColumnDefinitions.Count - 1; i++)
            {
                for (int j = 0; j < table.RowDefinitions.Count; j++)
                {
                    if(j == 0)
                    {
                        txt1 = GetValueGrid(j, i, table);
                    }
                    else
                    {
                        txt2 = GetValueGrid(j, i, table);
                    }                  
                }
                if(int.TryParse(txt1, out time))
                {
                    if(float.TryParse(txt2, out power))
                    {
                        dishargePairs.Add(new DishargePairs { dischTime = time, dischPower = power });
                    }
                    else
                    {
                        dishargePairs.Add(new DishargePairs { dischTime = 0, dischPower = 0 });
                        MessageBox.Show("Неверный формат введенных данных");
                    }
                }
                else
                {
                    dishargePairs.Add(new DishargePairs { dischTime = 0, dischPower = 0 });
                    MessageBox.Show("Неверный формат введенных данных");
                }              
            }
        }
        public static void DelPair()
        {
            dishargePairs.RemoveAt(dishargePairs.Count - 1);
        }
        private static string GetValueGrid(int row, int col, Grid table)//Функция получения значений из Grid по номеру ряда и столбца
        {
            string txt = "";
            if (row >= 0 && col >= 0)
            {
                TextBox ch = (table.Children.Cast<UIElement>().First(b => Grid.GetRow(b) == row && Grid.GetColumn(b) == col + 1) as TextBox);
                if(ch == null)
                {
                    txt = "0";
                }
                else
                {
                    txt = (ch.Text).ToString();
                    if (txt == string.Empty)
                        txt = "0";
                }
            }
            else txt = "0";
            return txt;
        }
    }
}
