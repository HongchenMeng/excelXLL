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
        /// 获取农历
        /// </summary>
        /// <param name="dtStr">时间字符串</param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取农历日期")]
        public static string GetLunarDate([ExcelArgument(Description = "时间")] string dtStr)
        { 
            string[] dd = dtStr.Split('/');
            int y = int.Parse(dd[0]);
            int m = int.Parse(dd[1]);
            int d = int.Parse(dd[2]);
            DateTime dt = new DateTime(y, m, d);
            ChinaDate cd = new ChinaDate(dt);

            return cd.nlDate;
        }

        /// <summary>
        /// 获取农历干支
        /// </summary>
        /// <param name="dtStr">时间字符串</param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取农历日期")]
        public static string GetGZ([ExcelArgument(Description = "时间")] string dtStr)
        {
            string[] dd = dtStr.Split('/');
            int y = int.Parse(dd[0]);
            int m = int.Parse(dd[1]);
            int d = int.Parse(dd[2]);
            DateTime dt = new DateTime(y, m, d);
            ChinaDate cd = new ChinaDate(dt);

            return cd.gzDate;
        }
        /// <summary>
        /// 获取公历日期
        /// </summary>
        /// <param name="lunarYear">农历年份</param>
        /// <param name="lunarMonth">农历月份</param>
        /// <param name="lunarDay">农历日</param>
        /// <param name="theMonthIsLeap">该月是否闰月</param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取公历日期")]
        public static string GetSolarDate([ExcelArgument(Description = "农历年份")] int lunarYear, [ExcelArgument(Description = "农历月份")] int lunarMonth, [ExcelArgument(Description = "农历日")] int lunarDay, [ExcelArgument(Description = "该月是否农历润月")] bool theMonthIsLeap)
        {
            DateTimeLunar dl = new DateTimeLunar();
            DateTime dt = dl.GetSolarDate(lunarYear, lunarMonth, lunarDay, theMonthIsLeap);

            return dt.ToShortDateString();
        }
        /// <summary>
        /// 获取农历每月的天数
        /// </summary>
        /// <param name="lunarYear">农历年</param>
        /// <param name="lunarMonth">农历月</param>
        /// <returns></returns>
        [ExcelFunction(Category = "test测试分类", IsMacroType = true, Description = "获取农历每月天数")]
        public static int GetDaysInLunarMonth([ExcelArgument(Description = "农历年份")] int lunarYear, [ExcelArgument(Description = "农历月份")] int lunarMonth)
        {
            DateTimeLunar dl = new DateTimeLunar();
            int days = dl.GetDaysInLunarMonth(lunarYear, lunarMonth);

            return days;
        }
    }
}
