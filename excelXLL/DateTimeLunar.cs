﻿using System;

namespace excelXLL
{
    internal class DateTimeLunar
    {
        #region 农历常量

        /// <summary>
        /// 公历与农历互转，公历最小时间
        /// </summary>
        private static DateTime MinDay = new DateTime(1900, 1, 30);

        /// <summary>
        /// 公历与农历互转，公历最大时间
        /// </summary>
        private static DateTime MaxDay = new DateTime(2099, 12, 31);

        /// <summary>
        /// 农历月份字符：正、二、三……十、冬、腊
        /// </summary>
        private static string[] MonthStr = { "正", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "九", "拾", "冬", "腊" };

        /// <summary>
        /// 农历日子字符，初、十、廿、卅
        /// </summary>
        private static string[] DayStr = { "初", "十", "廿", "卅" };

        /// <summary>
        /// 农历数据
        /// </summary>
        private static int[] LunarDateArray = new int[] {
     0x04AE53,0x0A5748,0x5526BD,0x0D2650,0x0D9544,
     0x46AAB9,0x056A4D,0x09AD42,0x24AEB6,0x04AE4A, //1901-1910

     0x6A4DBE,0x0A4D52,0x0D2546,0x5D52BA,0x0B544E,
     0x0D6A43,0x296D37,0x095B4B,0x749BC1,0x049754, //1911-1920

     0x0A4B48,0x5B25BC,0x06A550,0x06D445,0x4ADAB8,
     0x02B64D,0x095742,0x2497B7,0x04974A,0x664B3E, //1921-1930

     0x0D4A51,0x0EA546,0x56D4BA,0x05AD4E,0x02B644,
     0x393738,0x092E4B,0x7C96BF,0x0C9553,0x0D4A48, //1931-1940

     0x6DA53B,0x0B554F,0x056A45,0x4AADB9,0x025D4D,
     0x092D42,0x2C95B6,0x0A954A,0x7B4ABD,0x06CA51, //1941-1950

     0x0B5546,0x555ABB,0x04DA4E,0x0A5B43,0x352BB8,
     0x052B4C,0x8A953F,0x0E9552,0x06AA48,0x7AD53C, //1951-1960

     0x0AB54F,0x04B645,0x4A5739,0x0A574D,0x052642,
     0x3E9335,0x0D9549,0x75AABE,0x056A51,0x096D46, //1961-1970

     0x54AEBB,0x04AD4F,0x0A4D43,0x4D26B7,0x0D254B,
     0x8D52BF,0x0B5452,0x0B6A47,0x696D3C,0x095B50, //1971-1980

     0x049B45,0x4A4BB9,0x0A4B4D,0xAB25C2,0x06A554,
     0x06D449,0x6ADA3D,0x0AB651,0x093746,0x5497BB, //1981-1990

     0x04974F,0x064B44,0x36A537,0x0EA54A,0x86B2BF,
     0x05AC53,0x0AB647,0x5936BC,0x092E50,0x0C9645, //1991-2000

     0x4D4AB8,0x0D4A4C,0x0DA541,0x25AAB6,0x056A49,
     0x7AADBD,0x025D52,0x092D47,0x5C95BA,0x0A954E, //2001-2010

     0x0B4A43,0x4B5537,0x0AD54A,0x955ABF,0x04BA53,
     0x0A5B48,0x652BBC,0x052B50,0x0A9345,0x474AB9, //2011-2020

     0x06AA4C,0x0AD541,0x24DAB6,0x04B64A,0x69573D,
     0x0A4E51,0x0D2646,0x5E933A,0x0D534D,0x05AA43, //2021-2030

     0x36B537,0x096D4B,0xB4AEBF,0x04AD53,0x0A4D48,
     0x6D25BC,0x0D254F,0x0D5244,0x5DAA38,0x0B5A4C, //2031-2040

     0x056D41,0x24ADB6,0x049B4A,0x7A4BBE,0x0A4B51,
     0x0AA546,0x5B52BA,0x06D24E,0x0ADA42,0x355B37, //2041-2050

     0x09374B,0x8497C1,0x049753,0x064B48,0x66A53C,
     0x0EA54F,0x06B244,0x4AB638,0x0AAE4C,0x092E42, //2051-2060

     0x3C9735,0x0C9649,0x7D4ABD,0x0D4A51,0x0DA545,
     0x55AABA,0x056A4E,0x0A6D43,0x452EB7,0x052D4B, //2061-2070

     0x8A95BF,0x0A9553,0x0B4A47,0x6B553B,0x0AD54F,
     0x055A45,0x4A5D38,0x0A5B4C,0x052B42,0x3A93B6, //2071-2080

     0x069349,0x7729BD,0x06AA51,0x0AD546,0x54DABA,
     0x04B64E,0x0A5743,0x452738,0x0D264A,0x8E933E, //2081-2090

     0x0D5252,0x0DAA47,0x66B53B,0x056D4F,0x04AE45,
     0x4A4EB9,0x0A4D4C,0x0D1541,0x2D92B5 //2091-2099
        };

        #endregion 农历常量

        /// <summary>
        /// 公历每月第一天是公历年中的第几天（非润年）
        /// </summary>
        private static int[] NormalYday = { 1, 32, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335 };

        /// <summary>
        /// 公历每月第一天是公历年中的第几天（润年）
        /// </summary>
        private static int[] LeapYday = { 1, 32, 61, 92, 122, 153, 183, 214, 245, 275, 306, 336 };

        #region 私有变量

        /// <summary>
        /// 私有变量，能否公历转农历
        /// </summary>
        private bool _Solar2Lunar = false;

        /// <summary>
        /// 私有变量，该年的农历数据
        /// </summary>
        private int _LunarData;

        /// <summary>
        /// 私有变量，公历日期
        /// </summary>
        private DateTime _SolarDate;

        /// <summary>
        /// 私有变量，农历年份
        /// </summary>
        private int _LunarYear;

        /// <summary>
        /// 私有变量，农历月份
        /// </summary>
        private int _LunarMonth;

        /// <summary>
        /// 私有变量，农历日
        /// </summary>
        private int _LunarDay;

        /// <summary>
        /// 私有变量，公历年是否闰年
        /// </summary>
        private bool _IsLeapSolarYear;

        /// <summary>
        /// 私有变量，农历年是否闰年
        /// </summary>
        private bool _IsLeapLunarYear;

        /// <summary>
        /// 私有变量，农历年哪个月是润月，6代表润五月
        /// </summary>
        private int _WhereLeapLunarMonth;

        /// <summary>
        /// 私有变量，农历年1~13个月每个月的天数，闰月之后的月份数+1
        /// </summary>
        private int[] _DaysOfLunarMonth = { 29, 29, 29, 29, 29, 29, 29, 29, 29, 29, 29, 29, 0 };

        /// <summary>
        /// 私有变量，农历年全年共有多少天
        /// </summary>
        private int _DaysOfLunarYear = 348; //29天 X 12个月

        /// <summary>
        /// 私有变量，该公历年春节对应的公历日期
        /// </summary>
        public DateTime DateTimeOfSpringFestival;

        #endregion 私有变量

        #region 公共变量

        public bool Solar2Lunar
        {
            get
            {
                if (_SolarDate != DateTime.Parse("0001-01-01"))
                {
                    return _Solar2Lunar;
                }
                else
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// 公历日期
        /// </summary>
        public DateTime SolarDate
        {
            get
            {
                return _SolarDate;
            }
        }

        /// <summary>
        /// 农历年份
        /// </summary>
        public int LunarYear
        {
            set
            {
                _LunarYear = value;
            }
            get
            {
                if (this.SolarDate < this.DateTimeOfSpringFestival)
                {
                    return SolarDate.Year - 1;
                }
                else
                {
                    return SolarDate.Year;
                }
            }
        }

        /// <summary>
        /// 农历月份
        /// </summary>
        public int LunarMonth
        {
            set
            {
                _LunarMonth = value;
            }
            get
            {
                return _LunarMonth;
            }
        }

        /// <summary>
        /// 农历日
        /// </summary>
        public int LunarDay
        {
            set
            {
                _LunarDay = value;
            }
            get
            {
                return _LunarDay;
            }
        }

        /// <summary>
        /// 公历年是否闰年
        /// </summary>
        public bool IsLeapSolarYear
        {
            get
            {
                return _IsLeapSolarYear;
            }
        }

        /// <summary>
        /// 农历年是否闰年
        /// </summary>
        public bool IsLeapLunarYear
        {
            get
            {
                return _IsLeapLunarYear;
            }
        }

        /// <summary>
        /// 农历年润哪个月，6代表润六月
        /// </summary>
        public int WhereLeapLunarMonth
        {
            get
            {
                if (_WhereLeapLunarMonth > 0)
                {
                    return _WhereLeapLunarMonth - 1;
                }
                else
                {
                    return 0;
                }
            }
        }

        /// <summary>
        /// 农历年全年共有多少天
        /// </summary>
        public int DaysOfLunarYear
        {
            get
            {
                return _DaysOfLunarYear;
            }
        }

        #endregion 公共变量

        #region 构造方法

        public DateTimeLunar()
        {
        }

        public DateTimeLunar(DateTime datetime)
        {
            _SolarDate = datetime;
            Solar2LunarEnabled(_SolarDate);
        }

        #endregion 构造方法

        #region 私有方法

        /// <summary>
        /// 获取指定年份对应的农历数据
        /// </summary>
        /// <param name="year"></param>
        /// <returns></returns>
        private int GetLunarData(int year)
        {
            return LunarDateArray[year - MinDay.Year];
        }

        ///// <summary>
        ///// 获取春节对应的公历日期
        ///// </summary>
        //private void GetLuanarData()
        //{
        //    _LunarData = GetLunarData(_SolarDate.Year);
        //    //if (((OldData.LunarCalendarTable[year - 1901] & 0x0060) >> 5) == 1)
        //    //    month = 1;
        //    //else
        //    //    month = 2;
        //}

        #endregion 私有方法

        #region 公共方法

        /// <summary>
        /// 设置公历年能否转换农历年，公历年日期必须在最小转换日期和最大转换日期之间才能转换
        /// </summary>
        public void Solar2LunarEnabled(DateTime dt)
        {
            if (dt >= MinDay && dt <= MaxDay)
            {
                _Solar2Lunar = true;
            }
        }

        /// <summary>
        /// 返回指定日期时间春节所对应的公历日期
        /// </summary>
        /// <param name="dt">指定日期时间</param>
        /// <returns></returns>
        public DateTime GetSpringFestivalDate(DateTime dt)
        {
            int year = dt.Year;
            return GetSpringFestivalDate(year);
        }

        /// <summary>
        /// 返回指定年份春节所对应的公历日期
        /// </summary>
        /// <param name="year">年份</param>
        /// <returns></returns>
        public DateTime GetSpringFestivalDate(int year)
        {
            int lunarData = GetLunarData(year);
            int month;
            if (((lunarData & 0x0060) >> 5) == 1)//0x0060 转二进制为：1100000
            {
                month = 1;
            }
            else
            {
                month = 2;
            }
            int day = lunarData & 0x1F;

            DateTime tDateTime = new DateTime(year, month, day);
            return tDateTime;
        }

        /// <summary>
        /// 返回公历日期对应的农历年份
        /// </summary>
        /// <param name="dt">公历日期</param>
        /// <returns></returns>
        public int GetLunarYear(DateTime dt)
        {
            DateTime dtSpringFestival = GetSpringFestivalDate(dt);

            return GetLunarYear(dt, dtSpringFestival);
        }

        /// <summary>
        /// 返回公历日期对应的农历年份
        /// </summary>
        /// <param name="dt">公历日期</param>
        /// <param name="dtSpringFestival">公历年春节对应的公历日期</param>
        /// <returns></returns>
        public int GetLunarYear(DateTime dt, DateTime dtSpringFestival)
        {
            int year = dt.Year;
            if (dt < dtSpringFestival)
            {
                return year - 1;
            }
            else
            {
                return year;
            }
        }

        public string GetLunarString(DateTime dt)
        {
            int solarYear = dt.Year;
            int lunarYear = GetLunarYear(dt);
            int lunarMonth=1;
            int lunarDay;//农历日子
            int lunarData = GetLunarData(lunarYear);

            int dayOfLunarYear = DayOfLunarYear(dt);
            //int lunarMonth = 1;
            for (; lunarMonth <= 13; lunarMonth++)
            {
                int daysOfLunarMonth = 29;
                if (((lunarData >> (6 + lunarMonth)) & 0x1) == 1)//大月就30天
                {
                    daysOfLunarMonth = 30;
                }
                if (dayOfLunarYear <= daysOfLunarMonth)
                {
                    break;
                }
                else
                {
                    dayOfLunarYear -= daysOfLunarMonth;

                }
            }
            lunarDay = dayOfLunarYear;
            int leapMonth = (lunarData >> 20) & 0xf;
            if(leapMonth >0 && leapMonth <lunarMonth)
            {
                lunarMonth--;

            }
            //int leap_month = (LUNAR_YEARS[year_index] >> 20) & 0xf;
            //if (leap_month > 0 && leap_month < luanr_month)
            //{
            //    luanr_month--;
            //    //如果当前月为闰月，设置闰月标志

            //    if (luanr_month == leap_month)
            //        luanr_date.leap = true;
            //}

            return LunarYear + "-" + lunarMonth + "-" + LunarDay;
        }

        /// <summary>
        /// 返回指定日期所在年份是否公历闰年
        /// </summary>
        /// <param name="dt">指定日期</param>
        /// <returns></returns>
        public bool LeapSolarYear(DateTime dt)
        {
            int year = dt.Year;
            return LeapSolarYear(year);
        }

        /// <summary>
        /// 返回指定公历年份是否为公历闰年
        /// </summary>
        /// <param name="year">公历年份</param>
        /// <returns></returns>
        public bool LeapSolarYear(int year)
        {
            if (year % 400 == 0)
            {
                return true;
            }
            else if (year % 4 == 0 && year % 100 != 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 计算指定日期距离元旦几天（是公历年中的第几天）
        /// </summary>
        /// <param name="dt">公历日期</param>
        /// <returns></returns>
        public int DayOfSolarYear(DateTime dt)
        {
            int month = dt.Month;
            int day = dt.Day;

            int[] Yday;//每月1日是公历年的第几天
            if (LeapSolarYear(dt))
            {
                Yday = LeapYday;
            }
            else
            {
                Yday = NormalYday;
            }

            return Yday[month - 1] + day - 1;
        }

        /// <summary>
        /// 计算指定年份 春节距离元旦的天数
        /// </summary>
        /// <param name="year"></param>
        /// <returns></returns>
        public int DayOfSolarYear(int year)
        {
            DateTime dt = GetSpringFestivalDate(year);
            return DayOfSolarYear(dt);
        }

        /// <summary>
        /// 计算指定公历日期是农历年的第几天
        /// </summary>
        /// <param name="dt">公历日期</param>
        /// <returns></returns>
        public int DayOfLunarYear(DateTime dt)
        {
            int year = dt.Year;
            int d = DayOfSolarYear(dt) - DayOfSolarYear(year) + 1;
            if (d < 0)//农历年份比公历年份小
            {
                int lunardata = GetLunarData(year - 1);

                int sm = (lunardata & 0x60) >> 5;//春节月
                int sd = (lunardata & 0x1f);//春节日

                DateTime dt2 = new DateTime(year - 1, 12, 31);
                d = DayOfSolarYear(dt2) - DayOfSolarYear(dt2.Year) + 1 + DayOfSolarYear(dt);
            }
            return d;
        }

        #endregion 公共方法
    }
}