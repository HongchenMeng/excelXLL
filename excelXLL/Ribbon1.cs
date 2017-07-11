using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using ExcelDna.Integration.CustomUI;
using ExcelDnaIRibbons = ExcelDna.Integration.CustomUI.IRibbonControl;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using ExcelDna.IntelliSense;
//using Microsoft.Office.Core;
//using IRibbonC = Microsoft.Office.Core.IRibbonControl;
// TODO:   按照以下步骤启用功能区(XML)项: 

// 1. 将以下代码块复制到 ThisAddin、ThisWorkbook 或 ThisDocument 类中。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. 在此类的“功能区回调”区域中创建回调方法，以处理用户
//    操作(如单击某个按钮)。注意: 如果已经从功能区设计器中导出此功能区，
//    则将事件处理程序中的代码移动到回调方法并修改该代码以用于
//    功能区扩展性(RibbonX)编程模型。

// 3. 向功能区 XML 文件中的控制标记分配特性，以标识代码中的相应回调方法。  

// 有关详细信息，请参见 Visual Studio Tools for Office 帮助中的功能区 XML 文档。

namespace excelXLL
{
    /// <summary>
    /// 选项卡函数回调
    /// </summary>
    [ComVisible(true)]
    public class Ribbon1 : ExcelRibbon
    {
        /// <summary>
        /// Excel应用程序
        /// </summary>
        public  Microsoft.Office.Interop.Excel.Application xlApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;

        private Microsoft.Office.Tools.CustomTaskPaneCollection CustomTaskPanes;

        private ExcelDna.Integration.CustomUI.IRibbonUI ribbon;
        //private ExcelDnaIRibbons ribbon;
        //选项卡Tab是否显示
        private bool TabIconsbool;
        private bool TabAgurbool;
        private bool TabHomebool;
        private bool TabInsertbool;
        private bool TabPageLayoutExcelbool;
        private bool TabFormulasbool;
        private bool TabDatabool;
        private bool TabReviewbool;
        private bool TabViewbool;
        private bool TabDeveloperbool;
        private bool TabAddInsbool;

        //工作表导航窗格控制
        private bool togbtnTaskPaneOfSheetLeftbool;
        private bool togbtnTaskPaneOfSheetRightbool;
        private bool togbtnTaskPaneOfSheetFloatbool;

        //定义用户控件
        private NavigationOfSheet navigationOfsheet = null;
        //定义面板对象
        private Microsoft.Office.Tools.CustomTaskPane taskSheet = null;

        /// <summary>
        /// 选项卡初始化
        /// </summary>
        /// <param name="ribbonUI"></param>
        public void Ribbon_Load(ExcelDna.Integration.CustomUI.IRibbonUI ribbonUI)
        {
            //内置图标Tab
            TabIconsbool = false;
            TabAgurbool = false;
            TabHomebool = true;
            TabInsertbool = true;
            TabPageLayoutExcelbool = true;
            TabFormulasbool = true;
            TabDatabool = true;
            TabReviewbool = true;
            TabViewbool = true;
            TabDeveloperbool = true;
            TabAddInsbool = false;

            //工作表导航窗格控制
            togbtnTaskPaneOfSheetLeftbool = true;
            togbtnTaskPaneOfSheetRightbool = false;
            togbtnTaskPaneOfSheetFloatbool = false;

            this.ribbon = ribbonUI;

            //navigationOfsheet = new NavigationOfSheet();
            ////创建面板
            //taskSheet = CustomTaskPanes.Add(navigationOfsheet, "工作表导航");

            //// 设置面板停靠位置
            //taskSheet.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
            //// 设置面板是否可以被移动
            ////taskSheet.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            //// 绑定面板的可见性变化事件
            //taskSheet.VisibleChanged += taskSheet_VisibleChanged;

            //// 设置面板默认可见
            //taskSheet.Visible = true;
        }
        /// <summary>
        /// 自定义图标调用，没有的话自定义图标显示不出来
        /// </summary>
        /// <param name="ImageName"></param>
        /// <returns></returns>
        public override object LoadImage(string ImageName)
        {
            object obj = Resource1.ResourceManager.GetObject(ImageName);
            return ((System.Drawing.Bitmap)(obj));
        }

        /// <summary>
        /// 刷新选项卡显隐状态
        /// </summary>
        /// <param name="control">ExcelDNA中的选项卡</param>
        /// <returns></returns>
        public bool TabgetVisible(ExcelDna.Integration.CustomUI.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "TabHomeb":
                    return TabHomebool;
                case "TabInsert":
                    return TabInsertbool;
                case "TabPageLayoutExcel":
                    return TabPageLayoutExcelbool;
                case "TabFormulas":
                    return TabFormulasbool;
                case "TabData":
                    return TabDatabool;
                case "TabReview":
                    return TabReviewbool;
                case "TabView":
                    return TabViewbool;
                case "TabDeveloper":
                    return TabDeveloperbool;
                case "TabAddIns":
                    return TabAddInsbool;
                case "TabIcons":
                    return TabIconsbool;
                case "TabAgur":
                    return TabAgurbool;
                default:
                    return false;
            }
        }
        /// <summary>
        ///  动态控制选项卡显隐状态
        /// </summary>
        /// <param name="control">ExcelDNA中的选项卡</param>
        /// <param name="pressed">boot型</param>
        public void OAcheckBoxShowTab(ExcelDna.Integration.CustomUI.IRibbonControl control, bool pressed)
        {
            switch (control.Id)
            {
                case "checkBoxShowTabHomeb":
                    TabHomebool = pressed;
                    break;
                case "checkBoxShowTabInsert":
                    TabInsertbool = pressed;
                    break;
                case "checkBoxShowTabPageLayoutExcel":
                    TabPageLayoutExcelbool = pressed;
                    break;
                case "checkBoxShowTabFormulas":
                    TabFormulasbool = pressed;
                    break;
                case "checkBoxShowTabData":
                    TabDatabool = pressed;
                    break;
                case "checkBoxShowTabReview":
                    TabReviewbool = pressed;
                    break;
                case "checkBoxShowTabView":
                    TabViewbool = pressed;
                    break;
                case "checkBoxShowTabDeveloper":
                    TabDeveloperbool = pressed;
                    break;
                case "checkBoxShowTabAddIns":
                    TabAddInsbool = pressed;
                    break;
                case "checkBoxShowTabIcons":
                    TabIconsbool = pressed;
                    break;
                case "checkBoxShowTabAgur":
                    TabAgurbool = pressed;
                    break;
                default:
                    break;
            }
            this.ribbon.Invalidate();

        }
        /// <summary>
        /// 选项卡控制勾选标识
        /// </summary>
        /// <param name="control"></param>
        /// <returns></returns>
        public bool checkBoxShowTabgetPressed(ExcelDna.Integration.CustomUI.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "checkBoxShowTabHomeb":
                    return TabHomebool;
                case "checkBoxShowTabInsert":
                    return TabInsertbool;
                case "checkBoxShowTabPageLayoutExcel":
                    return TabPageLayoutExcelbool;
                case "checkBoxShowTabFormulas":
                    return TabFormulasbool;
                case "checkBoxShowTabData":
                    return TabDatabool;
                case "checkBoxShowTabReview":
                    return TabReviewbool;
                case "checkBoxShowTabView":
                    return TabViewbool;
                case "checkBoxShowTabDeveloper":
                    return TabDeveloperbool;
                case "checkBoxShowTabAddIns":
                    return TabAddInsbool;
                case "checkBoxShowTabIcons":
                    return TabIconsbool;
                case "checkBoxShowTabAgur":
                    return TabAgurbool;
                default:
                    return false;
            }
        }
        /// <summary>
        /// 显示内置图标回调
        /// </summary>
        /// <param name="control"></param>
        /// <param name="selectedId"></param>
        /// <param name="selectedIndex"></param>
        public void OAShowImageMso(ExcelDna.Integration.CustomUI.IRibbonControl control, string selectedId, int selectedIndex)//命名要与xml中一致
        {
            Microsoft.Office.Interop.Excel.Range ActiveCell = (Microsoft.Office.Interop.Excel.Range)xlApp.ActiveCell;
           
            ActiveCell.Value = selectedId;
        }

        #region 工作表导航窗格
        ///// <summary>
        ///// 面板可见性变化时，更新界面对应的ToggleButton状态
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void taskSheet_VisibleChanged(object sender, EventArgs e)
        //{
        //    this.ribbon.InvalidateControl("togbtnTaskPaneOfSheetLeft");
        //    this.ribbon.InvalidateControl("togbtnTaskPaneOfSheetRight");
        //    this.ribbon.InvalidateControl("togbtnTaskPaneOfSheetFloat");
        //}

        ///// <summary>
        ///// ToggleButton点击，切换面板可见性
        ///// </summary>
        ///// <param name="control"></param>
        ///// <param name="pressed"></param>
        //public void OAtogbtnTaskPaneOfSheet(IRibbonControl control, bool pressed)
        //{

        //    switch (control.Id)
        //    {
        //        case "togbtnTaskPaneOfSheetLeft":
        //            togbtnTaskPaneOfSheetLeftbool = !togbtnTaskPaneOfSheetLeftbool;
        //            if (togbtnTaskPaneOfSheetLeftbool == true)
        //            {
        //                togbtnTaskPaneOfSheetRightbool = false;
        //                togbtnTaskPaneOfSheetFloatbool = false;
        //            }

        //            taskSheet.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
        //            break;
        //        case "togbtnTaskPaneOfSheetRight":
        //            togbtnTaskPaneOfSheetRightbool = !togbtnTaskPaneOfSheetRightbool;
        //            if (togbtnTaskPaneOfSheetRightbool == true)
        //            {
        //                togbtnTaskPaneOfSheetLeftbool = false;
        //                togbtnTaskPaneOfSheetFloatbool = false;
        //            }
        //            taskSheet.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
        //            break;
        //        default:
        //            togbtnTaskPaneOfSheetFloatbool = !togbtnTaskPaneOfSheetFloatbool;
        //            if (togbtnTaskPaneOfSheetFloatbool == true)
        //            {
        //                togbtnTaskPaneOfSheetLeftbool = false;
        //                togbtnTaskPaneOfSheetRightbool = false;
        //            }
        //            taskSheet.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating;
        //            break;
        //    }
        //    taskSheet.Visible = pressed;//
        //    this.ribbon.InvalidateControl("togbtnTaskPaneOfSheetLeft");
        //    this.ribbon.InvalidateControl("togbtnTaskPaneOfSheetRight");
        //    this.ribbon.InvalidateControl("togbtnTaskPaneOfSheetFloat");

        //}

        ///// <summary>
        ///// 面板可见性的回掉函数。
        ///// </summary>
        ///// <param name="control"></param>
        ///// <returns></returns>
        //public bool getPressedbtnTaskPaneOfSheet(IRibbonControl control)
        //{
        //    // 按钮状态
        //    switch (control.Id)
        //    {
        //        case "togbtnTaskPaneOfSheetLeft":
        //            return togbtnTaskPaneOfSheetLeftbool;
        //        case "togbtnTaskPaneOfSheetRight":
        //            return togbtnTaskPaneOfSheetRightbool;
        //        default:
        //            return togbtnTaskPaneOfSheetFloatbool;
        //    }

        //}
        #endregion

    }
}
