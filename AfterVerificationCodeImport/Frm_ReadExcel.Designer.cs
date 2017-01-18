namespace AfterVerificationCodeImport
{
    partial class Frm_ReadExcel
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
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.xradioRegOnline = new DevExpress.XtraEditors.RadioGroup();
            this.btnPCT = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.cbxSheet = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lb_tablename = new System.Windows.Forms.Label();
            this.bt_ok = new System.Windows.Forms.Button();
            this.com_databasename = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dgv_show = new System.Windows.Forms.DataGridView();
            this.bt_see = new System.Windows.Forms.Button();
            this.txt_filepath = new System.Windows.Forms.TextBox();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.xradioRegOnline.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_show)).BeginInit();
            this.SuspendLayout();
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(668, 553);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(176, 23);
            this.button4.TabIndex = 56;
            this.button4.Text = "6.更新案件相关关系";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(668, 497);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(176, 23);
            this.button5.TabIndex = 54;
            this.button5.Text = "4.更新同族关系";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(668, 470);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(176, 23);
            this.button3.TabIndex = 53;
            this.button3.Text = "3.更新IDS关系";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(668, 413);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(176, 23);
            this.button2.TabIndex = 52;
            this.button2.Text = "1.更新数据";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(668, 524);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(176, 23);
            this.button1.TabIndex = 55;
            this.button1.Text = "5.更新优先权顺序";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.xradioRegOnline);
            this.groupBox1.Location = new System.Drawing.Point(17, 325);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(618, 178);
            this.groupBox1.TabIndex = 50;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "文件类型";
            // 
            // xradioRegOnline
            // 
            this.xradioRegOnline.EditValue = "=请选择=";
            this.xradioRegOnline.Location = new System.Drawing.Point(6, 20);
            this.xradioRegOnline.Name = "xradioRegOnline";
            this.xradioRegOnline.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
            this.xradioRegOnline.Properties.Appearance.Options.UseBackColor = true;
            this.xradioRegOnline.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
            new DevExpress.XtraEditors.Controls.RadioGroupItem("=请选择=", "=请选择="),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("年费", "年费"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("备注", "备注"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("优先权", "优先权"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("Update", "Update"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("客户信息", "客户信息"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("外代理信息", "外代理信息"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("任务时限", "任务时限"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("分案信息", "分案信息"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("档案位置", "档案位置"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("相关客户", "相关客户"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("要求配置", "要求配置"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("客户要求", "客户要求"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("部门核对", "部门核对"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("申请方式", "申请方式"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("要求/发文", "要求/发文"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("申请人信息", "申请人信息"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("总委托书号", "总委托书号"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("实体信息表", "实体信息表"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("自定义属性", "自定义属性"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("案件处理人", "案件处理人"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("相关客户要求", "相关客户要求"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("优先权（香港）", "优先权（香港）"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("相关案件-双申", "相关案件-双申"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("更新申请人译名", "更新申请人译名"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("澳门案件-澳门延伸", "澳门案件-澳门延伸"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("国内-收文数据导入", "国内-收文数据导入"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("国外-国外库发明人表", "国外-国外库发明人表"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("国内-OA数据补充导入", "国内-OA数据补充导入"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("国外-国外库时限备注表", "国外-国外库时限备注表"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("香港-专利数据补充导入", "香港-专利数据补充导入"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("国外-法律信息及日志表", "国外-法律信息及日志表"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("国内-专利数据补充导入", "国内-专利数据补充导入")});
            this.xradioRegOnline.Size = new System.Drawing.Size(606, 148);
            this.xradioRegOnline.TabIndex = 0;
            // 
            // btnPCT
            // 
            this.btnPCT.Location = new System.Drawing.Point(668, 443);
            this.btnPCT.Name = "btnPCT";
            this.btnPCT.Size = new System.Drawing.Size(176, 23);
            this.btnPCT.TabIndex = 49;
            this.btnPCT.Text = "2.更新PCT关系";
            this.btnPCT.UseVisualStyleBackColor = true;
            this.btnPCT.Click += new System.EventHandler(this.btnPCT_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(40, 512);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 12);
            this.label3.TabIndex = 48;
            this.label3.Text = "Sheet";
            // 
            // cbxSheet
            // 
            this.cbxSheet.FormattingEnabled = true;
            this.cbxSheet.Location = new System.Drawing.Point(83, 509);
            this.cbxSheet.Name = "cbxSheet";
            this.cbxSheet.Size = new System.Drawing.Size(183, 20);
            this.cbxSheet.TabIndex = 47;
            this.cbxSheet.SelectedIndexChanged += new System.EventHandler(this.cbxSheet_SelectedIndexChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(666, 368);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(125, 12);
            this.label9.TabIndex = 46;
            this.label9.Text = "2.请选择对应文件类型";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(666, 347);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(197, 12);
            this.label8.TabIndex = 45;
            this.label8.Text = "1.请选择所导入数据对应数据库名称";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(666, 391);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(173, 12);
            this.label6.TabIndex = 44;
            this.label6.Text = "3.请单击“确定导入”数据按钮";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(654, 325);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 12);
            this.label5.TabIndex = 43;
            this.label5.Text = "操作步骤：";
            // 
            // lb_tablename
            // 
            this.lb_tablename.AutoSize = true;
            this.lb_tablename.Location = new System.Drawing.Point(12, 292);
            this.lb_tablename.Name = "lb_tablename";
            this.lb_tablename.Size = new System.Drawing.Size(119, 12);
            this.lb_tablename.TabIndex = 42;
            this.lb_tablename.Text = "表名:(请选择数据源)";
            // 
            // bt_ok
            // 
            this.bt_ok.Location = new System.Drawing.Point(138, 548);
            this.bt_ok.Name = "bt_ok";
            this.bt_ok.Size = new System.Drawing.Size(75, 23);
            this.bt_ok.TabIndex = 41;
            this.bt_ok.Text = "确定导入";
            this.bt_ok.UseVisualStyleBackColor = true;
            this.bt_ok.Click += new System.EventHandler(this.bt_ok_Click);
            // 
            // com_databasename
            // 
            this.com_databasename.FormattingEnabled = true;
            this.com_databasename.Location = new System.Drawing.Point(351, 512);
            this.com_databasename.Name = "com_databasename";
            this.com_databasename.Size = new System.Drawing.Size(183, 20);
            this.com_databasename.TabIndex = 40;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(304, 515);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 39;
            this.label2.Text = "数据库";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 38;
            this.label1.Text = "数据源";
            // 
            // dgv_show
            // 
            this.dgv_show.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_show.Location = new System.Drawing.Point(9, 38);
            this.dgv_show.Name = "dgv_show";
            this.dgv_show.RowTemplate.Height = 23;
            this.dgv_show.Size = new System.Drawing.Size(877, 251);
            this.dgv_show.TabIndex = 37;
            // 
            // bt_see
            // 
            this.bt_see.Location = new System.Drawing.Point(373, 10);
            this.bt_see.Name = "bt_see";
            this.bt_see.Size = new System.Drawing.Size(75, 23);
            this.bt_see.TabIndex = 36;
            this.bt_see.Text = "浏览...";
            this.bt_see.UseVisualStyleBackColor = true;
            // 
            // txt_filepath
            // 
            this.txt_filepath.Location = new System.Drawing.Point(59, 11);
            this.txt_filepath.Name = "txt_filepath";
            this.txt_filepath.Size = new System.Drawing.Size(302, 21);
            this.txt_filepath.TabIndex = 34;
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(668, 582);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(176, 23);
            this.button6.TabIndex = 57;
            this.button6.Text = "7.相关客户要求";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(668, 611);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(176, 23);
            this.button7.TabIndex = 58;
            this.button7.Text = "8.更新客户官方收据抬头";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // Frm_ReadExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(946, 665);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnPCT);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cbxSheet);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.lb_tablename);
            this.Controls.Add(this.bt_ok);
            this.Controls.Add(this.com_databasename);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgv_show);
            this.Controls.Add(this.bt_see);
            this.Controls.Add(this.txt_filepath);
            this.Name = "Frm_ReadExcel";
            this.Text = "操作";
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.xradioRegOnline.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_show)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnPCT;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbxSheet;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lb_tablename;
        private System.Windows.Forms.Button bt_ok;
        private System.Windows.Forms.ComboBox com_databasename;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dgv_show;
        private System.Windows.Forms.Button bt_see;
        private System.Windows.Forms.TextBox txt_filepath;
        private DevExpress.XtraEditors.RadioGroup xradioRegOnline;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
    }
}