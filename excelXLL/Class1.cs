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
    public class Class1
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
    }
}
