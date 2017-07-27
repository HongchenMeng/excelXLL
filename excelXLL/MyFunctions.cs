using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Xml;
using ExcelDna.IntelliSense;
namespace excelXLL
{
    /// <summary>
    /// Excel自定义函数
    /// </summary>
    public class MyFunctions
    {
        /// <summary>
        /// 获取身份证号码的性别
        /// </summary>
        /// <param name="ID">身份证号码</param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description ="获取身份证号码的性别")]
        public static string GetIDSex([ExcelArgument(Description = "身份证号码")] string ID)
        {
            IDCardHelper idCardHelper = new IDCardHelper(ID);
            return idCardHelper.Sex;
        }
        /// <summary>
        /// 获取身份证号码的地址
        /// </summary>
        /// <param name="ID">身份证号码</param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取身份证号码的地址")]
        public static string GetIDArea([ExcelArgument( Description = "身份证号码")] string ID)
        {
            IDCardHelper idCardHelper = new IDCardHelper(ID);
            return idCardHelper.Area;
        }
        /// <summary>
        /// 获取身份证号码的出生日期
        /// </summary>
        /// <param name="ID">身份证号码</param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取身份证号码的出生日期")]
        public static string GetIDBirthday([ExcelArgument( Description = "身份证号码")] string ID)
        {
            IDCardHelper idCardHelper = new IDCardHelper(ID);
            DateTime dt = idCardHelper.Birthday;
            if(!idCardHelper.checkBoll)
            {
                return idCardHelper.CheckStr;
            }
            else if (dt > DateTime.Now)
            {
                return "穿越啦";
            }
            else if(dt !=DateTime.Parse("0001-01-01"))
            {
                return dt.ToShortDateString().ToString();
            }
            else
            {
                return "出生日期错误";
            }
            
        }
        /// <summary>
        /// 获取身份证号码的年龄
        /// </summary>
        /// <param name="ID">身份证号码</param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取身份证号码的年龄")]
        public static string GetIDAge([ExcelArgument(Description = "身份证号码")] string ID)
        {
            IDCardHelper idCardHelper = new IDCardHelper(ID);
            return idCardHelper.Age;
        }
        /// <summary>
        /// 获取干支
        /// </summary>
        /// <param name="dt1"></param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取干支")]
        public static string GetNL([ExcelArgument(Description = "日期")] string  dt1)
        {
            DateTime dt = DateTime.ParseExact(dt1, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None);
            CalendarInfo ci = new CalendarInfo(dt);
            //return ci.GanZhiAnimalYearString + ci.GanZhiMonthString + ci.GanZhiDateString;
            return ci.ChineseYearString + ci.ChineseMonthString + ci.ChineseDayString;
        }
        /// <summary>
        /// 获取润月月份
        /// </summary>
        /// <param name="year"></param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取润月月份")]
        public static int GetLeapMonth([ExcelArgument(Description = "年份")] int year)
        {
            System.Globalization.ChineseLunisolarCalendar cc;
            cc = new System.Globalization.ChineseLunisolarCalendar();
            int a= cc.GetLeapMonth(year);
            if (a > 0)
                --a;
            return a;
        }
        /// <summary>
        /// 获取月份天数
        /// </summary>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取月份天数")]
        public static int GetLeapDaysOfMonth([ExcelArgument(Description = "年份")] int year, [ExcelArgument(Description = "月份")] int month)
        {

            System.Globalization.ChineseLunisolarCalendar cc;
            cc = new System.Globalization.ChineseLunisolarCalendar();

            int a = cc.GetLeapMonth(year);
            if (a > 0)
                --a;

            if(a>0)
            {
                if(month>=a)
                {
                    return cc.GetDaysInMonth(year, month+1, 1);
                }
                else
                {
                    return cc.GetDaysInMonth(year, month, 1);
                }
            }
            else
            {
                return cc.GetDaysInMonth(year, month, 1);
            }
            
        }
        /// <summary>
        /// 获取月份天数
        /// </summary>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取月份1~13天数")]
        public static int GetLeapDaysOfMonth2([ExcelArgument(Description = "年份")] int year, [ExcelArgument(Description = "月份")] int month)
        {

            System.Globalization.ChineseLunisolarCalendar cc;
            cc = new System.Globalization.ChineseLunisolarCalendar();
            if (month >= 13 & cc.GetLeapMonth(year) <= 0)
                return 0;
                return cc.GetDaysInMonth(year, month, 1);

        }
        /// <summary>
        /// 获取润月天数
        /// </summary>
        /// <param name="year"></param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取润月天数")]
        public static int GetDaysOfLeapdMonth([ExcelArgument(Description = "年份")] int year)
        {
            System.Globalization.ChineseLunisolarCalendar cc;
            cc = new System.Globalization.ChineseLunisolarCalendar();
            int a = cc.GetLeapMonth(year);
            if(a>0)
            {
                return cc.GetDaysInMonth(year, a, 1);
            }
            else
            {
                return 0;
            }
        }
        /// <summary>
        /// 获取春节月份
        /// </summary>
        /// <param name="year"></param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取春节月份")]
        public static int GetMonth([ExcelArgument(Description = "年份")] int year)
        {
            System.Globalization.ChineseLunisolarCalendar cc;
            cc = new System.Globalization.ChineseLunisolarCalendar();

            return cc.ToDateTime(year, 1, 1, 0, 0, 0, 0).Month;
        }
        /// <summary>
        /// 获取春节日期
        /// </summary>
        /// <param name="year"></param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取春节日期")]
        public static int GetDate([ExcelArgument(Description = "年份")] int year)
        {
            System.Globalization.ChineseLunisolarCalendar cc;
            cc = new System.Globalization.ChineseLunisolarCalendar();

            return cc.ToDateTime(year, 1, 1, 0, 0, 0, 0).Day;
        }
        /// <summary>
        /// 获取农历
        /// </summary>
        /// <param name="dt1">公历字符串 yyyyMMdd</param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取农历日期")]
        public static string GetLunarDate([ExcelArgument(Description = "年份")] string dt1)
        {
            DateTime dt = DateTime.ParseExact(dt1, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None);
            DateTimeLunar dl = new DateTimeLunar();
            string st = dl.GetLunarString(dt);
            return st;
        }
        /// <summary>
        /// 获取农历
        /// </summary>
        /// <param name="dt1">公历字符串 yyyyMMdd</param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取农历日期")]
        public static string GetLunarDate2([ExcelArgument(Description = "年份")] string dt1)
        {
            DateTime dt = DateTime.ParseExact(dt1, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None);
            DateTimeLunar dl = new DateTimeLunar();
            dl.SolarDate = dt;

            return dl.GetLunarDate(LunarString.农历润六月初八);
        }
        /// <summary>
        /// 获取农历
        /// </summary>
        /// <param name="dt1">公历字符串 yyyyMMdd</param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取农历日期")]
        public static string GetLunarDate3([ExcelArgument(Description = "年份")] string dt1)
        {
            DateTime dt = DateTime.ParseExact(dt1, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None);
            DateTimeLunar dl = new DateTimeLunar();
            dl.SolarDate = dt;

            return dl.GetLunarDate(LunarString.农历润6月8日);
        }
        /// <summary>
        /// 获取农历
        /// </summary>
        /// <param name="dt1">公历字符串 yyyyMMdd</param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取农历日期")]
        public static string GetLunarDate4([ExcelArgument(Description = "年份")] string dt1)
        {
            DateTime dt = DateTime.ParseExact(dt1, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None);
            DateTimeLunar dl = new DateTimeLunar(dt);
            //dl.SolarDate = dt;

            return dl.GetLunarDate(LunarString.农历2017年润6月8日);
        }
        /// <summary>
        /// 获取农历
        /// </summary>
        /// <param name="dt1">公历字符串</param>
        /// <param name="ls">农历日期格式</param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取农历日期")]
        public static string GetLunarDate5([ExcelArgument(Description = "年份")] string dt1, [ExcelArgument(Description = "农历日期格式")] int ls)
        {
            DateTime dt = DateTime.ParseExact(dt1, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None);
            DateTimeLunar dl = new DateTimeLunar(dt);
            //dl.SolarDate = dt;

            return dl.GetLunarDate((LunarString)ls);
        }
    }
}
