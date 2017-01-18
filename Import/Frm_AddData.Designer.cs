namespace Import
{
    partial class Frm_AddData
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
            this.lb_tablename = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dgv_show = new System.Windows.Forms.DataGridView();
            this.bt_see = new System.Windows.Forms.Button();
            this.txt_filepath = new System.Windows.Forms.TextBox();
            this.bt_ok = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.cbxSheet = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_show)).BeginInit();
            this.SuspendLayout();
            // 
            // lb_tablename
            // 
            this.lb_tablename.AutoSize = true;
            this.lb_tablename.Location = new System.Drawing.Point(15, 292);
            this.lb_tablename.Name = "lb_tablename";
            this.lb_tablename.Size = new System.Drawing.Size(119, 12);
            this.lb_tablename.TabIndex = 15;
            this.lb_tablename.Text = "表名:(请选择数据源)";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 14;
            this.label1.Text = "数据源";
            // 
            // dgv_show
            // 
            this.dgv_show.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_show.Location = new System.Drawing.Point(12, 38);
            this.dgv_show.Name = "dgv_show";
            this.dgv_show.RowTemplate.Height = 23;
            this.dgv_show.Size = new System.Drawing.Size(877, 251);
            this.dgv_show.TabIndex = 13;
            // 
            // bt_see
            // 
            this.bt_see.Location = new System.Drawing.Point(376, 10);
            this.bt_see.Name = "bt_see";
            this.bt_see.Size = new System.Drawing.Size(75, 23);
            this.bt_see.TabIndex = 12;
            this.bt_see.Text = "浏览...";
            this.bt_see.UseVisualStyleBackColor = true;
            // 
            // txt_filepath
            // 
            this.txt_filepath.Location = new System.Drawing.Point(62, 11);
            this.txt_filepath.Name = "txt_filepath";
            this.txt_filepath.Size = new System.Drawing.Size(302, 21);
            this.txt_filepath.TabIndex = 11;
            // 
            // bt_ok
            // 
            this.bt_ok.Location = new System.Drawing.Point(92, 398);
            this.bt_ok.Name = "bt_ok";
            this.bt_ok.Size = new System.Drawing.Size(75, 23);
            this.bt_ok.TabIndex = 16;
            this.bt_ok.Text = "确定导入";
            this.bt_ok.UseVisualStyleBackColor = true;
            this.bt_ok.Click += new System.EventHandler(this.bt_ok_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(19, 323);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 12);
            this.label3.TabIndex = 25;
            this.label3.Text = "Sheet";
            // 
            // cbxSheet
            // 
            this.cbxSheet.FormattingEnabled = true;
            this.cbxSheet.Location = new System.Drawing.Point(62, 320);
            this.cbxSheet.Name = "cbxSheet";
            this.cbxSheet.Size = new System.Drawing.Size(183, 20);
            this.cbxSheet.TabIndex = 24;
            // 
            // Frm_AddData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(901, 692);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cbxSheet);
            this.Controls.Add(this.bt_ok);
            this.Controls.Add(this.lb_tablename);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgv_show);
            this.Controls.Add(this.bt_see);
            this.Controls.Add(this.txt_filepath);
            this.Name = "Frm_AddData";
            this.Text = "数据导入";
            ((System.ComponentModel.ISupportInitialize)(this.dgv_show)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lb_tablename;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dgv_show;
        private System.Windows.Forms.Button bt_see;
        private System.Windows.Forms.TextBox txt_filepath;
        private System.Windows.Forms.Button bt_ok;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbxSheet;
    }
}