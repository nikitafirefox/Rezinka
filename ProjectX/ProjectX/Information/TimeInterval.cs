using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace ProjectX.Information
{
    public class TimeInterval
    {
        public int Days { get; set; }
        public int Weeks { get; set; }
        public int Month { get; set; }

        public override string ToString()
        {
            string buf = "";
            if (Month > 0) {
                buf += Month + "мес.";
            }
            if (Weeks > 0) {
                buf += Weeks + "нед.";
            }

            if (Days > 0)
            {
                buf += Days + "дн.";
            }

            if (Days == 0 && Weeks == 0 && Month == 0) {

                buf = "в наличии";
            }

            return buf;
        }



        public TimeInterval(int days) {
            Days = days;
            Weeks = 0;
            Month = 0;
        }

        public TimeInterval()
        {
            Days = 0;
            Weeks = 0;
            Month = 0;
        }

        public TimeInterval(string str)
        {
            if (Regex.IsMatch(str, "[0-9]{1}дн\\."))
            {
                Days = int.Parse(Regex.Match(Regex.Match(str, "[0-9]{1}дн\\.").Value, "[0-9]{1}").Value);
            }
            else {
                Days = 0;
            }

            if (Regex.IsMatch(str, "[0-9]{1}нед\\."))
            {
                Weeks = int.Parse(Regex.Match(Regex.Match(str, "[0-9]{1}нед\\.").Value, "[0-9]{1}").Value);
            }
            else
            {
                Weeks = 0;
            }

            if (Regex.IsMatch(str, "[0-9]{1}мес\\."))
            {
                Month = int.Parse(Regex.Match(Regex.Match(str, "[0-9]{1}мес\\.").Value, "[0-9]{1}").Value);
            }
            else
            {
                Month = 0;
            }

        }

        public static bool operator <(TimeInterval time1, TimeInterval time2) {
            if (time1.Month == time2.Month)
            {
                if (time1.Weeks == time2.Weeks)
                {
                    return time1.Days < time2.Days;
                }
                else {
                    return time1.Weeks < time2.Weeks;
                }
            }
            else {
                return (time1.Month < time2.Month);
            }
        }

        public static bool operator >(TimeInterval time1, TimeInterval time2) {
            return time2 < time1;
        }

        public static bool operator <=(TimeInterval time1, TimeInterval time2)
        {
            if (time1.Month == time2.Month)
            {
                if (time1.Weeks == time2.Weeks)
                {
                    return time1.Days <= time2.Days;
                }
                else
                {
                    return time1.Weeks <= time2.Weeks;
                }
            }
            else
            {
                return (time1.Month <= time2.Month);
            }
        }

        public static bool operator >=(TimeInterval time1, TimeInterval time2)
        {
            return time2 <= time1;
        }

        public static bool operator ==(TimeInterval time1, TimeInterval time2) {
            return time1.Days == time2.Days && time1.Weeks == time2.Weeks && time1.Month == time2.Month;
        }

        public static bool operator !=(TimeInterval time1, TimeInterval time2) {
            return !(time1 == time2);
        }
    }
}
