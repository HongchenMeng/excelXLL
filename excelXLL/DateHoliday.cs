using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelXLL
{
    public partial class DateInfo
    {
        /// <summary>
        /// 公历假期结构，月/日/假期长度/假期名称
        /// </summary>
        private struct SolarHolidayStruct
        {
            public int Month;
            public int Day;
            /// <summary>
            /// 假期长度
            /// </summary>
            public int Recess; //假期长度
            /// <summary>
            /// 假期名称
            /// </summary>
            public string HolidayName;
            /// <summary>
            /// 公历假期
            /// </summary>
            /// <param name="month"></param>
            /// <param name="day"></param>
            /// <param name="recess"></param>
            /// <param name="name"></param>
            public SolarHolidayStruct(int month, int day, int recess, string name)
            {
                Month = month;
                Day = day;
                Recess = recess;
                HolidayName = name;
            }
        }

        /// <summary>
        /// 农历假期结构,月、日、假期长度、假期名称
        /// </summary>
        private struct LunarHolidayStruct
        {
            public int Month;
            public int Day;
            public int Recess;
            public string HolidayName;
            /// <summary>
            /// 农历节日，如(1, 1, 1, "春节")，(8, 15, 0, "中秋节")
            /// </summary>
            /// <param name="month">农历月</param>
            /// <param name="day">农历日</param>
            /// <param name="recess">假期长度</param>
            /// <param name="name">假期名称</param>
            public LunarHolidayStruct(int month, int day, int recess, string name)
            {
                Month = month;
                Day = day;
                Recess = recess;
                HolidayName = name;
            }
        }
        /// <summary>
        /// 按星期计算的假期结构，月、第几周、周几，假期名称
        /// </summary>
        private struct WeekHolidayStruct
        {
            public int Month;
            public int WeekAtMonth;
            public int WeekDay;
            public string HolidayName;
            /// <summary>
            /// 按星期计算的节日，如(5, 2, 1, "母亲节")五月第2个星期日
            /// </summary>
            /// <param name="month"></param>
            /// <param name="weekAtMonth"></param>
            /// <param name="weekDay"></param>
            /// <param name="name"></param>
            public WeekHolidayStruct(int month, int weekAtMonth, int weekDay, string name)
            {
                Month = month;
                WeekAtMonth = weekAtMonth;
                WeekDay = weekDay;
                HolidayName = name;
            }
        }
    }
}
