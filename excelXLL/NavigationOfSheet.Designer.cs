namespace excelXLL
{
    partial class NavigationOfSheet
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.listBoxOfSheet = new System.Windows.Forms.ListBox();
            this.buttonOfSheet = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // listBoxOfSheet
            // 
            this.listBoxOfSheet.FormattingEnabled = true;
            this.listBoxOfSheet.ItemHeight = 12;
            this.listBoxOfSheet.Location = new System.Drawing.Point(2, 8);
            this.listBoxOfSheet.Name = "listBoxOfSheet";
            this.listBoxOfSheet.Size = new System.Drawing.Size(177, 388);
            this.listBoxOfSheet.TabIndex = 0;
            this.listBoxOfSheet.SelectedIndexChanged += new System.EventHandler(this.listBoxOfSheet_SelectedIndexChanged);
            // 
            // buttonOfSheet
            // 
            this.buttonOfSheet.Location = new System.Drawing.Point(0, 414);
            this.buttonOfSheet.Name = "buttonOfSheet";
            this.buttonOfSheet.Size = new System.Drawing.Size(178, 39);
            this.buttonOfSheet.TabIndex = 1;
            this.buttonOfSheet.Text = "更新";
            this.buttonOfSheet.UseVisualStyleBackColor = true;
            this.buttonOfSheet.Click += new System.EventHandler(this.buttonOfSheet_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(39, 460);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(113, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "刷新显示工作表目录";
            // 
            // NavigationOfSheet
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonOfSheet);
            this.Controls.Add(this.listBoxOfSheet);
            this.Name = "NavigationOfSheet";
            this.Size = new System.Drawing.Size(180, 481);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listBoxOfSheet;
        private System.Windows.Forms.Button buttonOfSheet;
        private System.Windows.Forms.Label label1;
    }
}
