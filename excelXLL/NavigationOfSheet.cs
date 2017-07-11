using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace excelXLL
{
    /// <summary>
    /// 自定义控制面板
    /// </summary>
    public partial class NavigationOfSheet : UserControl
    {
        /// <summary>
        /// 自定义面板构造函数
        /// </summary>
        public NavigationOfSheet()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void buttonOfSheet_Click(object sender, EventArgs e)
        {
            this.listBoxOfSheet.Items.Clear();//事先清空列表框
            //foreach (Excel.Worksheet wst in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            //{
            //    this.listBoxOfSheet.Items.Add(wst.Name);//历遍工作表，加入表名
            //}
        }

        private void listBoxOfSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[this.listBoxOfSheet.Text].Activate();
        }
    }
}
