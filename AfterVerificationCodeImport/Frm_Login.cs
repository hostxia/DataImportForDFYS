using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AfterVerificationCodeImport
{
    public partial class Frm_Login : Form
    {
        DBConnection sqlconnstr;
        public Frm_Login()
        {
            InitializeComponent();
            sqlconnstr = new DBConnection();
            Load += new EventHandler(SetServer_Load);
        }
        void SetServer_Load(object sender, EventArgs e)
        {
            Init();
            Evenhand();
        }
        void Init()
        {
            AcceptButton = bt_Conn_Test;
            MaximizeBox = false;
            MaximumSize = MinimumSize = Size;
        }
        void Evenhand()
        {
            //单选按钮
            rd_SqlServer.Click += new EventHandler(rd_Click);
            rd_Windows.Click += new EventHandler(rd_Click);
            //按钮
            bt_Conn_Test.Click += new EventHandler(bt_Click);
            bt_Login.Click += new EventHandler(bt_Click);
        }
        void rd_Click(object sender, EventArgs e)
        {
            switch ((sender as RadioButton).Name)
            {
                case "rd_Windows":
                    lb_Uid.Enabled = false;
                    lb_Pwd.Enabled = false;
                    txt_Uid.Enabled = false;
                    txt_Pwd.Enabled = false;
                    break;
                case "rd_SqlServer":
                    lb_Uid.Enabled = true;
                    lb_Pwd.Enabled = true;
                    txt_Uid.Enabled = true;
                    txt_Pwd.Enabled = true;
                    break;
                default:
                    break;
            }
        }
        void bt_Click(object sender, EventArgs e)
        {
            if (txt_Server.Text.Trim().Length == 0)
            { MessageBox.Show("ServerName Is Null ."); return; }
            if (rd_SqlServer.Checked)
            {
                if (txt_Uid.Text.Trim().Length == 0)
                { MessageBox.Show("Uid Is Null ."); return; }
                if (txt_Pwd.Text.Trim().Length == 0)
                { MessageBox.Show("Pwd Is Null ."); return; }
            }
            switch ((sender as Button).Name)
            {
                case "bt_Conn_Test"://测试连接
                    bt_Login.Enabled = false;
                    sqlconnstr.ServerName = txt_Server.Text.Trim();
                    sqlconnstr.LoginName = txt_Uid.Text.Trim();
                    sqlconnstr.Password = txt_Pwd.Text.Trim();
                    sqlconnstr.Database = "master";
                    if (rd_SqlServer.Checked)//sqlserver登陆
                    {
                        using (SqlConnection conn = new SqlConnection(string.Format(@"server={0};database=master;uid={1};pwd={2}",
                            txt_Server.Text.Trim(), txt_Uid.Text.Trim(), txt_Pwd.Text.Trim())))
                        {
                            try
                            {
                                sqlconnstr.ConnectionString = conn.ConnectionString;
                                conn.Open();
                                bt_Login.Enabled = true;
                                AcceptButton = bt_Login;
                            }
                            catch
                            { MessageBox.Show("服务器连接失败!"); }
                        }
                    }
                    else//windows身份验证过
                    {
                        using (SqlConnection conn = new SqlConnection(string.Format(@"Data Source={0};database = master;Integrated security = true", txt_Server.Text.Trim())))
                        {
                            try
                            { sqlconnstr.ConnectionString = conn.ConnectionString; conn.Open(); bt_Login.Enabled = true; AcceptButton = bt_Login; }
                            catch
                            { MessageBox.Show("服务器连接失败!"); }
                        }
                    }
                    break;
                case "bt_Login"://登陆
                    Frm_ReadExcel show = new Frm_ReadExcel(sqlconnstr);
                    show.Show();
                    Hide();
                    break;
                default:
                    break;
            }
        }
    }
}
