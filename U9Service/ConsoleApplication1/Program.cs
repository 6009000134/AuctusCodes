using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine(Convert.ToInt32(DateTime.Now.DayOfWeek));
            Console.WriteLine(Convert.ToInt32(DateTime.Now.AddDays(1).DayOfWeek));
            Console.WriteLine(Convert.ToInt32(DateTime.Now.AddDays(2).DayOfWeek));
            Console.WriteLine(Convert.ToInt32(DateTime.Now.AddDays(3).DayOfWeek));
            Console.WriteLine(Convert.ToInt32(DateTime.Now.AddDays(4).DayOfWeek));
            Console.WriteLine(Convert.ToInt32(DateTime.Now.AddDays(5).DayOfWeek));
            Console.WriteLine(Convert.ToInt32(DateTime.Now.AddDays(6).DayOfWeek));
            Console.WriteLine(DateTime.Now.AddDays((Convert.ToInt32(DateTime.Now.DayOfWeek) - 1) * (-1)));
            for (int i = 0; i < 4; i++)
            {
                string d = Console.ReadLine();
                Console.WriteLine(GetMonday(d).ToShortDateString());

                string html = "";
                DateTime monday = GetMonday(d);
                for (int w = 0; w < 8; w++)
                {
                    html += "<td nowrap='nowrap'>" + monday.AddDays(w * 7).ToString("MM-dd") + "~" + monday.AddDays((w + 1) * 7 - 1).ToString("MM-dd");
                }
                Console.WriteLine(html);
            }

            Console.ReadLine();

        }
        public static DateTime GetMonday(string Date)
        {
            DateTime d = Convert.ToDateTime(Date);
            if (d.DayOfWeek == 0)
            {
                return d.AddDays(-6);
            }
            else
            {
                return d.AddDays((Convert.ToInt32(d.DayOfWeek) - 1) * (-1));
            }
        }
    }
}
