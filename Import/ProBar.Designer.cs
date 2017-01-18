namespace Import
{
    partial class ProBar
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.progressBarControl = new DevExpress.XtraEditors.ProgressBarControl();
            this.lbSuccess = new DevExpress.XtraEditors.LabelControl();
            this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
            this.lbTotalSelected = new DevExpress.XtraEditors.LabelControl();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
            this.lbTotalTime = new DevExpress.XtraEditors.LabelControl();
            this.lbTotalSum = new DevExpress.XtraEditors.LabelControl();
            this.labelControl5 = new DevExpress.XtraEditors.LabelControl();
            ((System.ComponentModel.ISupportInitialize)(this.progressBarControl.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // progressBarControl
            // 
            this.progressBarControl.Location = new System.Drawing.Point(30, 34);
            this.progressBarControl.Name = "progressBarControl";
            this.progressBarControl.Properties.ShowTitle = true;
            this.progressBarControl.Size = new System.Drawing.Size(562, 24);
            this.progressBarControl.TabIndex = 6;
            // 
            // lbSuccess
            // 
            this.lbSuccess.Location = new System.Drawing.Point(199, 104);
            this.lbSuccess.Name = "lbSuccess";
            this.lbSuccess.Size = new System.Drawing.Size(0, 14);
            this.lbSuccess.TabIndex = 10;
            // 
            // labelControl2
            // 
            this.labelControl2.Location = new System.Drawing.Point(132, 104);
            this.labelControl2.Name = "labelControl2";
            this.labelControl2.Size = new System.Drawing.Size(60, 14);
            this.labelControl2.TabIndex = 9;
            this.labelControl2.Text = "成功更新：";
            // 
            // lbTotalSelected
            // 
            this.lbTotalSelected.Location = new System.Drawing.Point(61, 104);
            this.lbTotalSelected.Name = "lbTotalSelected";
            this.lbTotalSelected.Size = new System.Drawing.Size(0, 14);
            this.lbTotalSelected.TabIndex = 8;
            // 
            // labelControl1
            // 
            this.labelControl1.Location = new System.Drawing.Point(30, 104);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(24, 14);
            this.labelControl1.TabIndex = 7;
            this.labelControl1.Text = "共：";
            // 
            // labelControl3
            // 
            this.labelControl3.Location = new System.Drawing.Point(340, 104);
            this.labelControl3.Name = "labelControl3";
            this.labelControl3.Size = new System.Drawing.Size(36, 14);
            this.labelControl3.TabIndex = 11;
            this.labelControl3.Text = "耗时：";
            // 
            // lbTotalTime
            // 
            this.lbTotalTime.Location = new System.Drawing.Point(382, 104);
            this.lbTotalTime.Name = "lbTotalTime";
            this.lbTotalTime.Size = new System.Drawing.Size(76, 14);
            this.lbTotalTime.TabIndex = 12;
            this.lbTotalTime.Text = "0天0时0分0秒";
            // 
            // lbTotalSum
            // 
            this.lbTotalSum.Location = new System.Drawing.Point(166, 139);
            this.lbTotalSum.Name = "lbTotalSum";
            this.lbTotalSum.Size = new System.Drawing.Size(7, 14);
            this.lbTotalSum.TabIndex = 14;
            this.lbTotalSum.Text = "0";
            // 
            // labelControl5
            // 
            this.labelControl5.Location = new System.Drawing.Point(30, 139);
            this.labelControl5.Name = "labelControl5";
            this.labelControl5.Size = new System.Drawing.Size(130, 14);
            this.labelControl5.TabIndex = 13;
            this.labelControl5.Text = "未找到\"我方卷号\"数量：";
            // 
            // ProBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(618, 209);
            this.Controls.Add(this.lbTotalSum);
            this.Controls.Add(this.labelControl5);
            this.Controls.Add(this.lbTotalTime);
            this.Controls.Add(this.labelControl3);
            this.Controls.Add(this.lbSuccess);
            this.Controls.Add(this.labelControl2);
            this.Controls.Add(this.lbTotalSelected);
            this.Controls.Add(this.labelControl1);
            this.Controls.Add(this.progressBarControl);
            this.Name = "ProBar";
            this.Text = "数据更新进度...";
            ((System.ComponentModel.ISupportInitialize)(this.progressBarControl.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public DevExpress.XtraEditors.ProgressBarControl progressBarControl;
        public DevExpress.XtraEditors.LabelControl lbSuccess;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        public DevExpress.XtraEditors.LabelControl lbTotalSelected;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraEditors.LabelControl labelControl3;
        public DevExpress.XtraEditors.LabelControl lbTotalTime;
        public DevExpress.XtraEditors.LabelControl lbTotalSum;
        private DevExpress.XtraEditors.LabelControl labelControl5;


    }
}