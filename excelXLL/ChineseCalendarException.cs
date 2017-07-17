using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelXLL
{
    /// <summary>
    /// 中国日历异常处理
    /// </summary>
    public class ChineseCalendarException : System.Exception
    {
        public ChineseCalendarException(string msg)
            : base(msg)
        {
        }
    }
}
