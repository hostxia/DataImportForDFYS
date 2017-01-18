namespace Winform_SqlBulkCopy
{
    partial class Frm_ReadExcel
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

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.txt_filepath = new System.Windows.Forms.TextBox();
            this.bt_see = new System.Windows.Forms.Button();
            this.dgv_show = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.com_databasename = new System.Windows.Forms.ComboBox();
            this.com_tablename = new System.Windows.Forms.ComboBox();
            this.bt_ok = new System.Windows.Forms.Button();
            this.lb_tablename = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.cbxSheet = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnPCT = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.xradioRegOnline = new DevExpress.XtraEditors.RadioGroup();
            this.label4 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_show)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.xradioRegOnline.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // txt_filepath
            // 
            this.txt_filepath.Location = new System.Drawing.Point(62, 12);
            this.txt_filepath.Name = "txt_filepath";
            this.txt_filepath.Size = new System.Drawing.Size(302, 21);
            this.txt_filepath.TabIndex = 0;
            // 
            // bt_see
            // 
            this.bt_see.Location = new System.Drawing.Point(376, 11);
            this.bt_see.Name = "bt_see";
            this.bt_see.Size = new System.Drawing.Size(75, 23);
            this.bt_see.TabIndex = 1;
            this.bt_see.Text = "浏览...";
            this.bt_see.UseVisualStyleBackColor = true;
            // 
            // dgv_show
            // 
            this.dgv_show.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_show.Location = new System.Drawing.Point(12, 39);
            this.dgv_show.Name = "dgv_show";
            this.dgv_show.RowTemplate.Height = 23;
            this.dgv_show.Size = new System.Drawing.Size(877, 251);
            this.dgv_show.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "数据源";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(307, 516);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "数据库";
            // 
            // com_databasename
            // 
            this.com_databasename.FormattingEnabled = true;
            this.com_databasename.Location = new System.Drawing.Point(354, 513);
            this.com_databasename.Name = "com_databasename";
            this.com_databasename.Size = new System.Drawing.Size(183, 20);
            this.com_databasename.TabIndex = 5;
            // 
            // com_tablename
            // 
            this.com_tablename.FormattingEnabled = true;
            this.com_tablename.Location = new System.Drawing.Point(354, 549);
            this.com_tablename.Name = "com_tablename";
            this.com_tablename.Size = new System.Drawing.Size(183, 20);
            this.com_tablename.TabIndex = 1;
            this.com_tablename.Visible = false;
            // 
            // bt_ok
            // 
            this.bt_ok.Location = new System.Drawing.Point(141, 549);
            this.bt_ok.Name = "bt_ok";
            this.bt_ok.Size = new System.Drawing.Size(75, 23);
            this.bt_ok.TabIndex = 8;
            this.bt_ok.Text = "确定导入";
            this.bt_ok.UseVisualStyleBackColor = true;
            // 
            // lb_tablename
            // 
            this.lb_tablename.AutoSize = true;
            this.lb_tablename.Location = new System.Drawing.Point(15, 293);
            this.lb_tablename.Name = "lb_tablename";
            this.lb_tablename.Size = new System.Drawing.Size(119, 12);
            this.lb_tablename.TabIndex = 10;
            this.lb_tablename.Text = "表名:(请选择数据源)";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(657, 326);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 12);
            this.label5.TabIndex = 15;
            this.label5.Text = "操作步骤：";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(669, 392);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(173, 12);
            this.label6.TabIndex = 16;
            this.label6.Text = "3.请单击“确定导入”数据按钮";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(669, 348);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(197, 12);
            this.label8.TabIndex = 19;
            this.label8.Text = "1.请选择所导入数据对应数据库名称";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(669, 369);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(125, 12);
            this.label9.TabIndex = 20;
            this.label9.Text = "2.请选择对应文件类型";
            // 
            // cbxSheet
            // 
            this.cbxSheet.FormattingEnabled = true;
            this.cbxSheet.Location = new System.Drawing.Point(86, 510);
            this.cbxSheet.Name = "cbxSheet";
            this.cbxSheet.Size = new System.Drawing.Size(183, 20);
            this.cbxSheet.TabIndex = 22;
            this.cbxSheet.SelectedIndexChanged += new System.EventHandler(this.cbxSheet_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(43, 513);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 12);
            this.label3.TabIndex = 23;
            this.label3.Text = "Sheet";
            // 
            // btnPCT
            // 
            this.btnPCT.Location = new System.Drawing.Point(671, 444);
            this.btnPCT.Name = "btnPCT";
            this.btnPCT.Size = new System.Drawing.Size(176, 23);
            this.btnPCT.TabIndex = 24;
            this.btnPCT.Text = "2.更新PCT关系";
            this.btnPCT.UseVisualStyleBackColor = true;
            this.btnPCT.Click += new System.EventHandler(this.btnPCT_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.xradioRegOnline);
            this.groupBox1.Location = new System.Drawing.Point(20, 326);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(618, 178);
            this.groupBox1.TabIndex = 26;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "文件类型";
            // 
            // xradioRegOnline
            // 
            this.xradioRegOnline.EditValue = "=请选择=";
            this.xradioRegOnline.Location = new System.Drawing.Point(24, 19);
            this.xradioRegOnline.Margin = new System.Windows.Forms.Padding(3, 10, 3, 3);
            this.xradioRegOnline.Name = "xradioRegOnline";
            this.xradioRegOnline.Properties.AllowMouseWheel = false;
            this.xradioRegOnline.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
            this.xradioRegOnline.Properties.Appearance.Options.UseBackColor = true;
            this.xradioRegOnline.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.xradioRegOnline.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
            new DevExpress.XtraEditors.Controls.RadioGroupItem("=请选择=", "=请选择="),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("年费", "年费"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("优先权", "优先权"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("优先权（香港）", "优先权（香港）"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("控制要求", "控制要求"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("相关客户", "相关客户"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("部门核对", "部门核对"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("任务时限", "任务时限"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("客户要求", "客户要求"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("递交机构", "递交机构"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("案件选择", "案件选择"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("申请方式", "申请方式"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("分案信息", "分案信息"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("总委托书号", "总委托书号"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("实体信息表", "实体信息表"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("案件处理人", "案件处理人"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("客户要求代码", "客户要求代码"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("相关案件-双申", "相关案件-双申"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("电子缴费单缴费人", "电子缴费单缴费人"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("澳门案件-澳门延伸", "澳门案件-澳门延伸"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("国内-收文数据导入", "国内-收文数据导入"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("国外-国外库发明人表", "国外-国外库发明人表"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("国内-OA数据补充导入", "国内-OA数据补充导入"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("国内-专利数据补充导入", "国内-专利数据补充导入"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("香港-专利数据补充导入", "香港-专利数据补充导入"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("国外-法律信息及日志表", "国外-法律信息及日志表"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("国外-国外库时限备注表", "国外-国外库时限备注表"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("客户信息", "客户信息"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("申请人信息", "申请人信息"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("外代理信息", "外代理信息"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("自定义属性", "自定义属性"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("要求/发文", "要求/发文"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("备注", "备注"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("档案位置", "档案位置"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("更新申请人译名", "更新申请人译名")});
            this.xradioRegOnline.Size = new System.Drawing.Size(588, 137);
            this.xradioRegOnline.TabIndex = 30;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(313, 552);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 12);
            this.label4.TabIndex = 27;
            this.label4.Text = "表名";
            this.label4.Visible = false;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(671, 414);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(176, 23);
            this.button2.TabIndex = 28;
            this.button2.Text = "1.更新数据";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(671, 471);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(176, 23);
            this.button3.TabIndex = 29;
            this.button3.Text = "3.更新IDS关系";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(671, 498);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(176, 23);
            this.button5.TabIndex = 31;
            this.button5.Text = "4.更新同族关系";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(671, 525);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(176, 23);
            this.button1.TabIndex = 32;
            this.button1.Text = "5.更新优先权顺序";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(671, 554);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(176, 23);
            this.button4.TabIndex = 33;
            this.button4.Text = "6.更新案件相关关系";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // Frm_ReadExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(907, 617);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label4);
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
            this.Controls.Add(this.com_tablename);
            this.Controls.Add(this.com_databasename);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgv_show);
            this.Controls.Add(this.bt_see);
            this.Controls.Add(this.txt_filepath);
            this.MaximizeBox = false;
            this.Name = "Frm_ReadExcel";
            this.Text = " 数据处理";
            ((System.ComponentModel.ISupportInitialize)(this.dgv_show)).EndInit();
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.xradioRegOnline.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txt_filepath;
        private System.Windows.Forms.Button bt_see;
        private System.Windows.Forms.DataGridView dgv_show;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox com_databasename;
        private System.Windows.Forms.ComboBox com_tablename;
        private System.Windows.Forms.Button bt_ok;
        private System.Windows.Forms.Label lb_tablename;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox cbxSheet;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnPCT;
        private System.Windows.Forms.GroupBox groupBox1;
        private DevExpress.XtraEditors.RadioGroup xradioRegOnline;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button4;



    }
}

