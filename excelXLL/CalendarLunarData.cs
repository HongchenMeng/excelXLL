using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelXLL
{
    public partial class CalendarInfo
    {
        /// <summary>
        /// 常量，农历计算最小年份
        /// </summary>
        private const int MinYear = 1900;
        /// <summary>
        /// 常量，农历计算最大年份
        /// </summary>
        private const int MaxYear = 2050;
        /// <summary>
        /// 农历最小年份初始化 1990.01.30
        /// </summary>
        private static DateTime MinDay = new DateTime(1900, 1, 30);
        /// <summary>
        /// 农历最大年份初始化 2049.12.31
        /// </summary>
        private static DateTime MaxDay = new DateTime(2049, 12, 31);


        /// <summary>
        /// 干支计算起始年
        /// </summary>
        private const int GanZhiStartYear = 1864; //干支计算起始年
        /// <summary>
        /// (1899, 12, 22);//起始日
        /// </summary>
        private static DateTime GanZhiStartDay = new DateTime(1899, 12, 22);//起始日
        private const string HZNum = "零一二三四五六七八九";
        /// <summary>
        /// 1900年为鼠年
        /// </summary>
        private const int AnimalStartYear = 1900; //1900年为鼠年
        /// <summary>
        /// (2007, 9, 13);//28星宿参考值,本日为角
        /// </summary>
        private static DateTime ChineseConstellationReferDay = new DateTime(2007, 9, 13);//28星宿参考值,本日为角

        /// <summary>
        /// 来源于网上的农历数据
        /// </summary>
        /// <remarks>
        /// 数据结构如下，共使用17位数据
        /// 第17位：表示闰月天数，0表示29天   1表示30天
        /// 第16位-第5位（共12位）表示12个月，其中第16位表示第一月，如果该月为30天则为1，29天为0
        /// 第4位-第1位（共4位）表示闰月是哪个月，如果当年没有闰月，则置0
        ///</remarks>
        private static int[] LunarDateArray = new int[]{
                0x04BD8,0x04AE0,0x0A570,0x054D5,0x0D260,0x0D950,0x16554,0x056A0,0x09AD0,0x055D2,//1900-1909
                0x04AE0,0x0A5B6,0x0A4D0,0x0D250,0x1D255,0x0B540,0x0D6A0,0x0ADA2,0x095B0,0x14977,//1910-1919
                0x04970,0x0A4B0,0x0B4B5,0x06A50,0x06D40,0x1AB54,0x02B60,0x09570,0x052F2,0x04970,//1920-1929
                0x06566,0x0D4A0,0x0EA50,0x06E95,0x05AD0,0x02B60,0x186E3,0x092E0,0x1C8D7,0x0C950,//1930-1929
                0x0D4A0,0x1D8A6,0x0B550,0x056A0,0x1A5B4,0x025D0,0x092D0,0x0D2B2,0x0A950,0x0B557,//1940-1949
                0x06CA0,0x0B550,0x15355,0x04DA0,0x0A5B0,0x14573,0x052B0,0x0A9A8,0x0E950,0x06AA0,//1950-1959
                0x0AEA6,0x0AB50,0x04B60,0x0AAE4,0x0A570,0x05260,0x0F263,0x0D950,0x05B57,0x056A0,//1960-1969
                0x096D0,0x04DD5,0x04AD0,0x0A4D0,0x0D4D4,0x0D250,0x0D558,0x0B540,0x0B6A0,0x195A6,//1970-1979
                0x095B0,0x049B0,0x0A974,0x0A4B0,0x0B27A,0x06A50,0x06D40,0x0AF46,0x0AB60,0x09570,//1980-1989
                0x04AF5,0x04970,0x064B0,0x074A3,0x0EA50,0x06B58,0x055C0,0x0AB60,0x096D5,0x092E0,//1990-1999
                0x0C960,0x0D954,0x0D4A0,0x0DA50,0x07552,0x056A0,0x0ABB7,0x025D0,0x092D0,0x0CAB5,//2000-2009
                0x0A950,0x0B4A0,0x0BAA4,0x0AD50,0x055D9,0x04BA0,0x0A5B0,0x15176,0x052B0,0x0A930,//2010-2019
                0x07954,0x06AA0,0x0AD50,0x05B52,0x04B60,0x0A6E6,0x0A4E0,0x0D260,0x0EA65,0x0D530,//2020-2029
                0x05AA0,0x076A3,0x096D0,0x04BD7,0x04AD0,0x0A4D0,0x1D0B6,0x0D250,0x0D520,0x0DD45,//2030-2039
                0x0B5A0,0x056D0,0x055B2,0x049B0,0x0A577,0x0A4B0,0x0AA50,0x1B255,0x06D20,0x0ADA0,//2040-2049
                0x14B63
                };

        /// <summary>
        /// 二十四节气:
        /// "小寒", "大寒", "立春", "雨水"
        /// "惊蛰", "春分", "清明", "谷雨"
        /// "立夏", "小满", "芒种", "夏至"
        /// "小暑", "大暑", "立秋", "处暑"
        /// "白露", "秋分", "寒露", "霜降"
        ///  "立冬", "小雪", "大雪", "冬至"
        /// </summary>
        private static string[] _lunarHolidayName =
                    {
                    "小寒", "大寒", "立春", "雨水",
                    "惊蛰", "春分", "清明", "谷雨",
                    "立夏", "小满", "芒种", "夏至",
                    "小暑", "大暑", "立秋", "处暑",
                    "白露", "秋分", "寒露", "霜降",
                    "立冬", "小雪", "大雪", "冬至"
                    };
        /// <summary>
        /// 二十八星宿
        /// "角木蛟","亢金龙","女土蝠","房日兔","心月狐","尾火虎","箕水豹"
        /// "斗木獬","牛金牛","氐土貉","虚日鼠","危月燕","室火猪","壁水獝"
        /// "奎木狼","娄金狗","胃土彘","昴日鸡","毕月乌","觜火猴","参水猿"
        ///  "井木犴","鬼金羊","柳土獐","星日马","张月鹿","翼火蛇","轸水蚓" 
        /// </summary>
        private static string[] _chineseConstellationName =
            {
                  //四        五      六         日        一      二      三  
                "角木蛟","亢金龙","女土蝠","房日兔","心月狐","尾火虎","箕水豹",
                "斗木獬","牛金牛","氐土貉","虚日鼠","危月燕","室火猪","壁水獝",
                "奎木狼","娄金狗","胃土彘","昴日鸡","毕月乌","觜火猴","参水猿",
                "井木犴","鬼金羊","柳土獐","星日马","张月鹿","翼火蛇","轸水蚓"
            };

        /// <summary>
        ///廿四节气名称
        /// </summary>
        private static string[] SolarTerm = new string[] { "小寒", "大寒", "立春", "雨水", "惊蛰", "春分", "清明", "谷雨", "立夏", "小满", "芒种", "夏至", "小暑", "大暑", "立秋", "处暑", "白露", "秋分", "寒露", "霜降", "立冬", "小雪", "大雪", "冬至" };
        /// <summary>
        /// 廿四节气对应数据
        /// </summary>
        private static int[] sTermInfo = new int[] { 0, 21208, 42467, 63836, 85337, 107014, 128867, 150921, 173149, 195551, 218072, 240693, 263343, 285989, 308563, 331033, 353350, 375494, 397447, 419210, 440795, 462224, 483532, 504758 };
        /// <summary>
        /// 干：甲乙丙丁戊己庚辛壬癸
        /// </summary>
        private static string ganStr = "甲乙丙丁戊己庚辛壬癸";
        /// <summary>
        /// 支：子丑寅卯辰巳午未申酉戌亥
        /// </summary>
        private static string zhiStr = "子丑寅卯辰巳午未申酉戌亥";
        /// <summary>
        /// 生肖：鼠牛虎兔龙蛇马羊猴鸡狗猪
        /// </summary>
        private static string animalStr = "鼠牛虎兔龙蛇马羊猴鸡狗猪";
        /// <summary>
        /// 日一二三四五六七八九
        /// </summary>
        private static string nStr1 = "日一二三四五六七八九";
        /// <summary>
        /// 初十廿卅
        /// </summary>
        private static string nStr2 = "初十廿卅";
        /// <summary>
        /// 月份："出错","正月","二月","三月"……
        /// </summary>
        private static string[] _monthString =
                {
                    "出错","正月","二月","三月","四月","五月","六月","七月","八月","九月","十月","十一月","腊月"
                };
        /// <summary>
        /// 按农历计算的节日
        /// </summary>
        private static LunarHolidayStruct[] lHolidayInfo = new LunarHolidayStruct[]{
            new LunarHolidayStruct(1, 1, 1, "春节"),
            new LunarHolidayStruct(1, 15, 0, "元宵节"),
            new LunarHolidayStruct(5, 5, 0, "端午节"),
            new LunarHolidayStruct(7, 7, 0, "七夕节"),
            new LunarHolidayStruct(7, 15, 0, "中元节"),
            new LunarHolidayStruct(8, 15, 0, "中秋节"),
            new LunarHolidayStruct(9, 9, 0, "重阳节"),
            new LunarHolidayStruct(12, 8, 0, "腊八节"),
            new LunarHolidayStruct(12, 23, 0, "扫房"),
            new LunarHolidayStruct(12, 24, 0, "小年"),
            //new LunarHolidayStruct(12, 30, 0, "除夕")  //注意除夕需要其它方法进行计算
        };

    }
}
