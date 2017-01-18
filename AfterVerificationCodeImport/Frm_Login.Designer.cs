namespace AfterVerificationCodeImport
{
    partial class Frm_Login
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
            this.gb_Conn = new System.Windows.Forms.GroupBox();
            this.txt_Pwd = new System.Windows.Forms.TextBox();
            this.txt_Uid = new System.Windows.Forms.TextBox();
            this.txt_Server = new System.Windows.Forms.TextBox();
            this.lb_Pwd = new System.Windows.Forms.Label();
            this.lb_Uid = new System.Windows.Forms.Label();
            this.lb_Server = new System.Windows.Forms.Label();
            this.rd_SqlServer = new System.Windows.Forms.RadioButton();
            this.bt_Conn_Test = new System.Windows.Forms.Button();
            this.rd_Windows = new System.Windows.Forms.RadioButton();
            this.gb_Style = new System.Windows.Forms.GroupBox();
            this.bt_Login = new System.Windows.Forms.Button();
            this.gb_Conn.SuspendLayout();
            this.gb_Style.SuspendLayout();
            this.SuspendLayout();
            // 
            // gb_Conn
            // 
            this.gb_Conn.Controls.Add(this.txt_Pwd);
            this.gb_Conn.Controls.Add(this.txt_Uid);
            this.gb_Conn.Controls.Add(this.txt_Server);
            this.gb_Conn.Controls.Add(this.lb_Pwd);
            this.gb_Conn.Controls.Add(this.lb_Uid);
            this.gb_Conn.Controls.Add(this.lb_Server);
            this.gb_Conn.Location = new System.Drawing.Point(12, 12);
            this.gb_Conn.Name = "gb_Conn";
            this.gb_Conn.Size = new System.Drawing.Size(253, 116);
            this.gb_Conn.TabIndex = 8;
            this.gb_Conn.TabStop = false;
            this.gb_Conn.Text = "连接信息";
            // 
            // txt_Pwd
            // 
            this.txt_Pwd.Location = new System.Drawing.Point(96, 80);
            this.txt_Pwd.Name = "txt_Pwd";
            this.txt_Pwd.PasswordChar = '*';
            this.txt_Pwd.Size = new System.Drawing.Size(133, 21);
            this.txt_Pwd.TabIndex = 7;
            this.txt_Pwd.Tag = "BizSolution2IP#";
            this.txt_Pwd.Text = "123456";
            // 
            // txt_Uid
            // 
            this.txt_Uid.Location = new System.Drawing.Point(96, 50);
            this.txt_Uid.Name = "txt_Uid";
            this.txt_Uid.Size = new System.Drawing.Size(133, 21);
            this.txt_Uid.TabIndex = 6;
            this.txt_Uid.Text = "sa";
            // 
            // txt_Server
            // 
            this.txt_Server.Location = new System.Drawing.Point(96, 22);
            this.txt_Server.Name = "txt_Server";
            this.txt_Server.Size = new System.Drawing.Size(133, 21);
            this.txt_Server.TabIndex = 4;
            this.txt_Server.Text = "CHENZHIWEN\\SQLEXPRESS";
            // 
            // lb_Pwd
            // 
            this.lb_Pwd.AutoSize = true;
            this.lb_Pwd.Location = new System.Drawing.Point(43, 83);
            this.lb_Pwd.Name = "lb_Pwd";
            this.lb_Pwd.Size = new System.Drawing.Size(35, 12);
            this.lb_Pwd.TabIndex = 3;
            this.lb_Pwd.Text = "密码:";
            // 
            // lb_Uid
            // 
            this.lb_Uid.AutoSize = true;
            this.lb_Uid.Location = new System.Drawing.Point(43, 53);
            this.lb_Uid.Name = "lb_Uid";
            this.lb_Uid.Size = new System.Drawing.Size(47, 12);
            this.lb_Uid.TabIndex = 2;
            this.lb_Uid.Text = "用户名:";
            // 
            // lb_Server
            // 
            this.lb_Server.AutoSize = true;
            this.lb_Server.Location = new System.Drawing.Point(19, 25);
            this.lb_Server.Name = "lb_Server";
            this.lb_Server.Size = new System.Drawing.Size(71, 12);
            this.lb_Server.TabIndex = 0;
            this.lb_Server.Text = "服务器名称:";
            // 
            // rd_SqlServer
            // 
            this.rd_SqlServer.AutoSize = true;
            this.rd_SqlServer.Checked = true;
            this.rd_SqlServer.Location = new System.Drawing.Point(8, 28);
            this.rd_SqlServer.Name = "rd_SqlServer";
            this.rd_SqlServer.Size = new System.Drawing.Size(131, 16);
            this.rd_SqlServer.TabIndex = 0;
            this.rd_SqlServer.TabStop = true;
            this.rd_SqlServer.Text = "Sql Server身份验证";
            this.rd_SqlServer.UseVisualStyleBackColor = true;
            // 
            // bt_Conn_Test
            // 
            this.bt_Conn_Test.Location = new System.Drawing.Point(181, 155);
            this.bt_Conn_Test.Name = "bt_Conn_Test";
            this.bt_Conn_Test.Size = new System.Drawing.Size(75, 23);
            this.bt_Conn_Test.TabIndex = 10;
            this.bt_Conn_Test.Text = "测试连接";
            this.bt_Conn_Test.UseVisualStyleBackColor = true;
            // 
            // rd_Windows
            // 
            this.rd_Windows.AutoSize = true;
            this.rd_Windows.Location = new System.Drawing.Point(8, 59);
            this.rd_Windows.Name = "rd_Windows";
            this.rd_Windows.Size = new System.Drawing.Size(113, 16);
            this.rd_Windows.TabIndex = 1;
            this.rd_Windows.TabStop = true;
            this.rd_Windows.Text = "Windows身份验证";
            this.rd_Windows.UseVisualStyleBackColor = true;
            // 
            // gb_Style
            // 
            this.gb_Style.Controls.Add(this.rd_Windows);
            this.gb_Style.Controls.Add(this.rd_SqlServer);
            this.gb_Style.Location = new System.Drawing.Point(12, 134);
            this.gb_Style.Name = "gb_Style";
            this.gb_Style.Size = new System.Drawing.Size(146, 97);
            this.gb_Style.TabIndex = 9;
            this.gb_Style.TabStop = false;
            this.gb_Style.Text = "登陆方式";
            // 
            // bt_Login
            // 
            this.bt_Login.Enabled = false;
            this.bt_Login.Location = new System.Drawing.Point(181, 190);
            this.bt_Login.Name = "bt_Login";
            this.bt_Login.Size = new System.Drawing.Size(75, 23);
            this.bt_Login.TabIndex = 11;
            this.bt_Login.Text = "登陆";
            this.bt_Login.UseVisualStyleBackColor = true;
            // 
            // Frm_Login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(290, 249);
            this.Controls.Add(this.gb_Conn);
            this.Controls.Add(this.bt_Conn_Test);
            this.Controls.Add(this.gb_Style);
            this.Controls.Add(this.bt_Login);
            this.Name = "Frm_Login";
            this.Text = "登录";
            this.gb_Conn.ResumeLayout(false);
            this.gb_Conn.PerformLayout();
            this.gb_Style.ResumeLayout(false);
            this.gb_Style.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gb_Conn;
        private System.Windows.Forms.TextBox txt_Pwd;
        private System.Windows.Forms.TextBox txt_Uid;
        private System.Windows.Forms.TextBox txt_Server;
        private System.Windows.Forms.Label lb_Pwd;
        private System.Windows.Forms.Label lb_Uid;
        private System.Windows.Forms.Label lb_Server;
        private System.Windows.Forms.RadioButton rd_SqlServer;
        private System.Windows.Forms.Button bt_Conn_Test;
        private System.Windows.Forms.RadioButton rd_Windows;
        private System.Windows.Forms.GroupBox gb_Style;
        private System.Windows.Forms.Button bt_Login;
    }
}

