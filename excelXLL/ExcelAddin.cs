using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.IntelliSense;

namespace excelXLL
{
    /// <summary>
    /// 启动时注册IntelliSense函数，实现自定义函数参数提示
    /// </summary>
    public class ExcelAddin:IExcelAddIn
    {
        /// <summary>
        /// 启动时注册IntelliSense函数，实现自定义函数参数提示
        /// 调试时会报错忽略即可
        /// </summary>
        public void AutoOpen()
        {
            IntelliSenseServer.Register();
            //ExcelDna.Logging.LogDisplay.Show();
            //ExcelDna.Logging.LogDisplay.DisplayOrder = ExcelDna.Logging.DisplayOrder.NewestFirst;
        }

        public void AutoClose()
        {
            // CONSIDER: Do we implement an explicit call here, or is the AppDomain Unload event good enough?
        }
    }
}
