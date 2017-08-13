using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using ExcelDnaIRibbons = ExcelDna.Integration.CustomUI.IRibbonControl;
using System.Windows.Forms;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.IntelliSense;
using System.Diagnostics;
using System.Drawing;

namespace excelXLL
{
   public class ExcelClass
    {

        private const int LOGPIXELSX = 88; //沿屏幕宽度每逻辑英寸的像素数，在多显示器系统中，该值对所显示器相同；
        private const int LOGPIXELSY = 90;//沿屏幕高度每逻辑英寸的像素数，在多显示器系统中，该值对所显示器相同；
        private const int TWIPSPERINCH = 1440;
        private int dpiX;//显示器像素
        private int dpiY;
        /// <summary>
        /// 获取当前Excel应用程序
        /// </summary>
        public Excel.Application ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
        private Excel.Workbook _ActWorkBook;
        private Excel.Range _ActCell;
        private Excel.Range _ActCellOffset;
        private POINT _ActCellPoint1 = new POINT(); //定义一个坐标结构
        private POINT _ActCellPoint2 = new POINT(); //定义一个坐标结构
        //private RECT _rect = new RECT();

        private string _ActWorkBookName;

        private IntPtr _hWnd;

        private Graphics g;
        public ExcelClass()
        {
            _ActWorkBook = ExcelApp.ActiveWorkbook;
            _ActWorkBookName = _ActWorkBook.Name;
             _ActCell = ExcelApp.ActiveCell;
            
            _ActCellOffset = _ActCell.Offset[1, 1];

            _ActCellPoint1.X = ExcelApp.ActiveWindow.PointsToScreenPixelsX(0)+Point2Pixel(_ActCell.Left,LOGPIXELSX);
            _ActCellPoint1.Y = ExcelApp.ActiveWindow.PointsToScreenPixelsY(0)+Point2Pixel(_ActCell.Top,LOGPIXELSY);
            
            _ActCellPoint2.X = ExcelApp.ActiveWindow.PointsToScreenPixelsX(0)+Point2Pixel(_ActCellOffset.Left,LOGPIXELSX);
            _ActCellPoint2.Y = ExcelApp.ActiveWindow.PointsToScreenPixelsY(0)+Point2Pixel(_ActCellOffset.Top,LOGPIXELSY);
            // _hWnd= WindowFromPoint(_ActCellPoint);
            IntPtr hd1 = FindWindow("XLMAIN", _ActWorkBookName + " - Excel");
            IntPtr hd2 = FindWindowEx(hd1, IntPtr.Zero, "XLDESK", null);//"XLDESK"
            _hWnd = FindWindowEx(hd2, IntPtr.Zero, "EXCEL7", _ActWorkBookName);// "EXCEL7"

            ScreenToClient(_hWnd,ref _ActCellPoint1);
            int x1 = _ActCellPoint1.X;
            int y1 = _ActCellPoint1.Y;

            ScreenToClient(_hWnd, ref _ActCellPoint2);
            int x2 = _ActCellPoint2.X;
            int y2 = _ActCellPoint2.Y;
            // _ActCellPoint=_hWnd.ToPointer()

            g = Graphics.FromHwnd(_hWnd);
            g.DrawRectangle(new Pen(Color.Red, 3), x1,y1, x2-x1,y2- y1);
          //  g.DrawRectangle(new Pen(Color.Red, 3), _rect);
          //  g.DrawRectangle(new Pen(Color.Yellow, 5), x2, y2, 10, 50);
            g.Dispose();



        }
        /// <summary>
        /// 获取窗口句柄
        /// </summary>
        /// <param name="lpClassName">窗口类名</param>
        /// <param name="lpWindowName">窗口标题</param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        /// <summary>
        /// 在窗口列表中寻找与指定条件相符的第一个子窗口 
        /// </summary>
        /// <param name="hWndParent">要查找的子窗口所在的父窗口的句柄（如果设置了hwndParent，则表示从这个hwndParent指向的父窗口中搜索子窗口）。如果hwndParent为 0 ，则函数以桌面窗口为父窗口，查找桌面窗口的所有子窗口。</param>
        /// <param name="hChildAfter">子窗口句柄。查找从在Z序中的下一个子窗口开始。子窗口必须为hwndParent窗口的直接子窗口而非后代窗口。如果HwndChildAfter为NULL，查找从hwndParent的第一个子窗口开始。如果hwndParent 和 hwndChildAfter同时为NULL，则函数查找所有的顶层窗口及消息窗口。</param>
        /// <param name="lpszClass">指向一个指定了类名的空结束字符串，或一个标识类名字符串的成员的指针。如果该参数为一个成员，则它必须为前次调用theGlobaIAddAtom函数产生的全局成员。该成员为16位，必须位于lpClassName的低16位，高位必须为0。</param>
        /// <param name="lpszWindowText">指向一个指定了窗口名（窗口标题）的空结束字符串。如果该参数为 NULL，则为所有窗口全匹配。</param>
        /// <returns>Long，找到的窗口的句柄。如未找到相符窗口，则返回零。会设置GetLastError.如果函数成功，返回值为具有指定类名和窗口名的窗口句柄。如果函数失败，返回值为NULL。</returns>
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hWndParent, IntPtr hChildAfter, string lpszClass, string lpszWindowText);
        [DllImport("user32", EntryPoint = "ScreenToClient")]
        public static extern bool ScreenToClient(IntPtr hwnd,ref POINT lpPoint);

        /// <summary>
        /// 获得包含指定点的窗口的句柄。
        /// HWND WindowFromPoint（POINT Point）；
        /// WindowFromPoint函数不获取隐藏或禁止的窗口句柄，即使点在该窗口内。应用程序应该使用ChildWindowFromPoint函数进行无限制查询，这样就可以获得静态文本控件的句柄。
        /// </summary>
        /// <param name="pt">指定一个被检测的点的POINT结构。</param>
        /// <returns>返回值为包含该点的窗口的句柄。如果包含指定点的窗口不存在，返回值为NULL。如果该点在静态文本控件之上，返回值是在该静态文本控件的下面的窗口的句柄。</returns>
        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(POINT pt);

        /// <summary>
        /// 获取指定设备的性能参数该方法将所取得的硬件设备信息保存到一个D3DCAPS9结构中。
        /// </summary>
        /// <param name="hwnd">句柄</param>
        /// <param name="nIndex">根据GetDeviceCaps索引表所示常数确定返回信息的类型</param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern int GetDeviceCaps(IntPtr hwnd, int nIndex);
        /// <summary>
        /// 获取桌面句柄
        /// </summary>
        /// <param name="hWnd">句柄</param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hWnd);
            /// <summary>
            /// 释放设备上下文环境（DC）供其他应用程序使用
            /// </summary>
            /// <param name="hWnd">指向要释放的设备上下文环境所在的窗口的句柄。</param>
            /// <param name="hdc">指向要释放的设备上下文环境的句柄</param>
            /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern bool ReleaseDC(IntPtr hWnd, IntPtr hdc);
        /// <summary>
        /// 获取显示器像素
        /// </summary>
        /// <param name="xy"></param>
        /// <param name="LOGPIXELSxy"></param>
        /// <returns></returns>
        public int Point2Pixel(int xy,int LOGPIXELSxy)
        {
            IntPtr inP = GetDC(IntPtr.Zero);

            int xy1 = GetDeviceCaps(inP, LOGPIXELSxy);
            int xy2 = xy / LOGPIXELSxy * xy1;
            ReleaseDC(IntPtr.Zero, inP);

            return xy2;

        }
        public string ActWorkBookName
        {
            get
            {
                return _ActWorkBookName;
            }
        }
        public IntPtr hWnd
        {
            get
            {
                return _hWnd;
            }
        }
        public void tset()
        {


           
        }
        public struct POINT
        {
            public int X;
            public int Y;
        }

        /// <summary>
        /// rect结构定义了一个矩形框左上角以及右下角的坐标,并绘制该矩形
        /// left ： 指定矩形框左上角的x坐标
        ///top： 指定矩形框左上角的y坐标
        ///right： 指定矩形框右下角的x坐标
        ///bottom：指定矩形框右下角的y坐标
        /// </summary>
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
            //public System.Drawing.Rectangle ToRectangle()
            //{
            //    return new System.Drawing.Rectangle(Left, Top, Right - Left, Bottom - Top);
            //}
        }

    }
}
