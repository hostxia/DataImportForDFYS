namespace AfterVerificationCodeImport
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
            this.lbTotalSum = new DevExpress.XtraEditors.LabelControl();
            this.labelControl5 = new DevExpress.XtraEditors.LabelControl();
            this.lbTotalTime = new DevExpress.XtraEditors.LabelControl();
            this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
            this.lbSuccess = new DevExpress.XtraEditors.LabelControl();
            this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
            this.lbTotalSelected = new DevExpress.XtraEditors.LabelControl();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.progressBarControl = new DevExpress.XtraEditors.ProgressBarControl();
            ((System.ComponentModel.ISupportInitialize)(this.progressBarControl.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // lbTotalSum
            // 
            this.lbTotalSum.Location = new System.Drawing.Point(91, 136);
            this.lbTotalSum.Name = "lbTotalSum";
            this.lbTotalSum.Size = new System.Drawing.Size(7, 14);
            this.lbTotalSum.TabIndex = 23;
            this.lbTotalSum.Text = "0";
            // 
            // labelControl5
            // 
            this.labelControl5.Location = new System.Drawing.Point(25, 136);
            this.labelControl5.Name = "labelControl5";
            this.labelControl5.Size = new System.Drawing.Size(60, 14);
            this.labelControl5.TabIndex = 22;
            this.labelControl5.Text = "错误数量：";
            // 
            // lbTotalTime
            // 
            this.lbTotalTime.Location = new System.Drawing.Point(377, 101);
            this.lbTotalTime.Name = "lbTotalTime";
            this.lbTotalTime.Size = new System.Drawing.Size(76, 14);
            this.lbTotalTime.TabIndex = 21;
            this.lbTotalTime.Text = "0天0时0分0秒";
            // 
            // labelControl3
            // 
            this.labelControl3.Location = new System.Drawing.Point(335, 101);
            this.labelControl3.Name = "labelControl3";
            this.labelControl3.Size = new System.Drawing.Size(36, 14);
            this.labelControl3.TabIndex = 20;
            this.labelControl3.Text = "耗时：";
            // 
            // lbSuccess
            // 
            this.lbSuccess.Location = new System.Drawing.Point(194, 101);
            this.lbSuccess.Name = "lbSuccess";
            this.lbSuccess.Size = new System.Drawing.Size(0, 14);
            this.lbSuccess.TabIndex = 19;
            // 
            // labelControl2
            // 
            this.labelControl2.Location = new System.Drawing.Point(127, 101);
            this.labelControl2.Name = "labelControl2";
            this.labelControl2.Size = new System.Drawing.Size(60, 14);
            this.labelControl2.TabIndex = 18;
            this.labelControl2.Text = "成功更新：";
            // 
            // lbTotalSelected
            // 
            this.lbTotalSelected.Location = new System.Drawing.Point(56, 101);
            this.lbTotalSelected.Name = "lbTotalSelected";
            this.lbTotalSelected.Size = new System.Drawing.Size(0, 14);
            this.lbTotalSelected.TabIndex = 17;
            // 
            // labelControl1
            // 
            this.labelControl1.Location = new System.Drawing.Point(25, 101);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(24, 14);
            this.labelControl1.TabIndex = 16;
            this.labelControl1.Text = "共：";
            // 
            // progressBarControl
            // 
            this.progressBarControl.Location = new System.Drawing.Point(25, 31);
            this.progressBarControl.Name = "progressBarControl";
            this.progressBarControl.Properties.ShowTitle = true;
            this.progressBarControl.Size = new System.Drawing.Size(562, 24);
            this.progressBarControl.TabIndex = 15;
            // 
            // ProBar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(619, 181);
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
            this.Text = "数据更新中......";
            ((System.ComponentModel.ISupportInitialize)(this.progressBarControl.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public DevExpress.XtraEditors.LabelControl lbTotalSum;
        private DevExpress.XtraEditors.LabelControl labelControl5;
        public DevExpress.XtraEditors.LabelControl lbTotalTime;
        private DevExpress.XtraEditors.LabelControl labelControl3;
        public DevExpress.XtraEditors.LabelControl lbSuccess;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        public DevExpress.XtraEditors.LabelControl lbTotalSelected;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        public DevExpress.XtraEditors.ProgressBarControl progressBarControl;
    }
}