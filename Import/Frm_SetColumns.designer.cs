namespace Winform_SqlBulkCopy
{
    partial class Frm_SetColumns
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.listbox_Excel = new System.Windows.Forms.ListBox();
            this.lb_Excel_First = new System.Windows.Forms.Label();
            this.lb_Excel_Before = new System.Windows.Forms.Label();
            this.lb_Excel_Last = new System.Windows.Forms.Label();
            this.lb_Excel_Next = new System.Windows.Forms.Label();
            this.lb_Sql_Last = new System.Windows.Forms.Label();
            this.lb_Sql_Next = new System.Windows.Forms.Label();
            this.lb_Sql_Before = new System.Windows.Forms.Label();
            this.lb_Sql_First = new System.Windows.Forms.Label();
            this.listbox_Ssms = new System.Windows.Forms.ListBox();
            this.bt_OK = new System.Windows.Forms.Button();
            this.lb_Excel_Del = new System.Windows.Forms.Label();
            this.lb_DB_Del = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(62, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "表格中字段";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(225, 28);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "数据库中字段";
            // 
            // listbox_Excel
            // 
            this.listbox_Excel.FormattingEnabled = true;
            this.listbox_Excel.ItemHeight = 12;
            this.listbox_Excel.Location = new System.Drawing.Point(28, 57);
            this.listbox_Excel.Name = "listbox_Excel";
            this.listbox_Excel.Size = new System.Drawing.Size(137, 208);
            this.listbox_Excel.TabIndex = 2;
            // 
            // lb_Excel_First
            // 
            this.lb_Excel_First.AutoSize = true;
            this.lb_Excel_First.Location = new System.Drawing.Point(171, 76);
            this.lb_Excel_First.Name = "lb_Excel_First";
            this.lb_Excel_First.Size = new System.Drawing.Size(17, 24);
            this.lb_Excel_First.TabIndex = 3;
            this.lb_Excel_First.Text = "一\r\n↑";
            // 
            // lb_Excel_Before
            // 
            this.lb_Excel_Before.AutoSize = true;
            this.lb_Excel_Before.Location = new System.Drawing.Point(171, 118);
            this.lb_Excel_Before.Name = "lb_Excel_Before";
            this.lb_Excel_Before.Size = new System.Drawing.Size(17, 12);
            this.lb_Excel_Before.TabIndex = 4;
            this.lb_Excel_Before.Text = "↑";
            // 
            // lb_Excel_Last
            // 
            this.lb_Excel_Last.AutoSize = true;
            this.lb_Excel_Last.Location = new System.Drawing.Point(171, 198);
            this.lb_Excel_Last.Name = "lb_Excel_Last";
            this.lb_Excel_Last.Size = new System.Drawing.Size(17, 24);
            this.lb_Excel_Last.TabIndex = 6;
            this.lb_Excel_Last.Text = "↓\r\n一";
            // 
            // lb_Excel_Next
            // 
            this.lb_Excel_Next.AutoSize = true;
            this.lb_Excel_Next.Location = new System.Drawing.Point(171, 156);
            this.lb_Excel_Next.Name = "lb_Excel_Next";
            this.lb_Excel_Next.Size = new System.Drawing.Size(17, 12);
            this.lb_Excel_Next.TabIndex = 5;
            this.lb_Excel_Next.Text = "↓";
            // 
            // lb_Sql_Last
            // 
            this.lb_Sql_Last.AutoSize = true;
            this.lb_Sql_Last.Location = new System.Drawing.Point(355, 198);
            this.lb_Sql_Last.Name = "lb_Sql_Last";
            this.lb_Sql_Last.Size = new System.Drawing.Size(17, 24);
            this.lb_Sql_Last.TabIndex = 11;
            this.lb_Sql_Last.Text = "↓\r\n一";
            // 
            // lb_Sql_Next
            // 
            this.lb_Sql_Next.AutoSize = true;
            this.lb_Sql_Next.Location = new System.Drawing.Point(355, 156);
            this.lb_Sql_Next.Name = "lb_Sql_Next";
            this.lb_Sql_Next.Size = new System.Drawing.Size(17, 12);
            this.lb_Sql_Next.TabIndex = 10;
            this.lb_Sql_Next.Text = "↓";
            // 
            // lb_Sql_Before
            // 
            this.lb_Sql_Before.AutoSize = true;
            this.lb_Sql_Before.Location = new System.Drawing.Point(355, 118);
            this.lb_Sql_Before.Name = "lb_Sql_Before";
            this.lb_Sql_Before.Size = new System.Drawing.Size(17, 12);
            this.lb_Sql_Before.TabIndex = 9;
            this.lb_Sql_Before.Text = "↑";
            // 
            // lb_Sql_First
            // 
            this.lb_Sql_First.AutoSize = true;
            this.lb_Sql_First.Location = new System.Drawing.Point(355, 76);
            this.lb_Sql_First.Name = "lb_Sql_First";
            this.lb_Sql_First.Size = new System.Drawing.Size(17, 24);
            this.lb_Sql_First.TabIndex = 8;
            this.lb_Sql_First.Text = "一\r\n↑";
            // 
            // listbox_Ssms
            // 
            this.listbox_Ssms.FormattingEnabled = true;
            this.listbox_Ssms.ItemHeight = 12;
            this.listbox_Ssms.Location = new System.Drawing.Point(212, 57);
            this.listbox_Ssms.Name = "listbox_Ssms";
            this.listbox_Ssms.Size = new System.Drawing.Size(137, 208);
            this.listbox_Ssms.TabIndex = 7;
            // 
            // bt_OK
            // 
            this.bt_OK.Location = new System.Drawing.Point(150, 287);
            this.bt_OK.Name = "bt_OK";
            this.bt_OK.Size = new System.Drawing.Size(75, 23);
            this.bt_OK.TabIndex = 12;
            this.bt_OK.Text = "设置完成";
            this.bt_OK.UseVisualStyleBackColor = true;
            // 
            // lb_Excel_Del
            // 
            this.lb_Excel_Del.AutoSize = true;
            this.lb_Excel_Del.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_Excel_Del.ForeColor = System.Drawing.Color.Red;
            this.lb_Excel_Del.Location = new System.Drawing.Point(171, 241);
            this.lb_Excel_Del.Name = "lb_Excel_Del";
            this.lb_Excel_Del.Size = new System.Drawing.Size(15, 15);
            this.lb_Excel_Del.TabIndex = 13;
            this.lb_Excel_Del.Text = "×";
            // 
            // lb_DB_Del
            // 
            this.lb_DB_Del.AutoSize = true;
            this.lb_DB_Del.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_DB_Del.ForeColor = System.Drawing.Color.Red;
            this.lb_DB_Del.Location = new System.Drawing.Point(357, 241);
            this.lb_DB_Del.Name = "lb_DB_Del";
            this.lb_DB_Del.Size = new System.Drawing.Size(15, 15);
            this.lb_DB_Del.TabIndex = 14;
            this.lb_DB_Del.Text = "×";
            // 
            // Frm_SetColumns
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(393, 351);
            this.Controls.Add(this.lb_DB_Del);
            this.Controls.Add(this.lb_Excel_Del);
            this.Controls.Add(this.bt_OK);
            this.Controls.Add(this.lb_Sql_Last);
            this.Controls.Add(this.lb_Sql_Next);
            this.Controls.Add(this.lb_Sql_Before);
            this.Controls.Add(this.lb_Sql_First);
            this.Controls.Add(this.listbox_Ssms);
            this.Controls.Add(this.lb_Excel_Last);
            this.Controls.Add(this.lb_Excel_Next);
            this.Controls.Add(this.lb_Excel_Before);
            this.Controls.Add(this.lb_Excel_First);
            this.Controls.Add(this.listbox_Excel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "Frm_SetColumns";
            this.Text = "设置列";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox listbox_Excel;
        private System.Windows.Forms.Label lb_Excel_First;
        private System.Windows.Forms.Label lb_Excel_Before;
        private System.Windows.Forms.Label lb_Excel_Last;
        private System.Windows.Forms.Label lb_Excel_Next;
        private System.Windows.Forms.Label lb_Sql_Last;
        private System.Windows.Forms.Label lb_Sql_Next;
        private System.Windows.Forms.Label lb_Sql_Before;
        private System.Windows.Forms.Label lb_Sql_First;
        private System.Windows.Forms.ListBox listbox_Ssms;
        private System.Windows.Forms.Button bt_OK;
        private System.Windows.Forms.Label lb_Excel_Del;
        private System.Windows.Forms.Label lb_DB_Del;
    }
}