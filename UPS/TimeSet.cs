using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UPS
{
    class TimeSet
    {
        public int time { get; set; }
        public string timStr { get; set; }
        public TimeSet(int time)
        {
            this.time = time;
            timStr = time.ToString() + " min";
        }
    }
}
