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
            System.Globalization.ChineseLunisolarCalendar cc;
            cc = new System.Globalization.ChineseLunisolarCalendar();

            Console.WriteLine(cc.GetLeapMonth(2017));//返回某年润月的月份，润六月 ，返回7
            Console.WriteLine(cc.GetDaysInYear(2014, 1));//返回某农历年全年的天数，第二个参数固定为1
            Console.WriteLine(cc.GetDaysInMonth(2017,6,1));//返回农历月全月的天数
            Console.WriteLine(cc.GetDaysInMonth(2017, 13, 1));//
            Console.WriteLine(cc.GetYear(DateTime.ParseExact("20170127", "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture)));//返回日期对应农历的年份2016
            Console.WriteLine(cc.GetYear(DateTime.ParseExact("20170128", "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture)));//返回日期对应农历的年份2017
            Console.WriteLine(cc.GetMonth(DateTime.ParseExact("20170723","yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture)));//返回日期对应农历的月份，1~13，有闰月需要额外处理
            Console.WriteLine(cc.GetDayOfMonth(DateTime.ParseExact("20170127", "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture)));//返回日期对应农历的日期
            Console.WriteLine(cc.ToDateTime(2017,1,1,0,0,0,0));//返回农历日期对应的公历日期
            

            Console.ReadLine();
        }
    }
}
