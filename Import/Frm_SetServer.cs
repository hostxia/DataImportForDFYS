using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Winform_SqlBulkCopy
{
    public partial class Frm_SetServer : Form
    {
        #region 全局变量
        /// <summary>
        /// 数据库连接字符串
        /// </summary>
        string sqlconnstr;
        #endregion

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        public Frm_SetServer()
        {
            InitializeComponent();
            Load += new EventHandler(SetServer_Load);
        }
        #endregion

        #region 窗体加载
        /// <summary>
        /// 窗体加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void SetServer_Load(object sender, EventArgs e)
        {
            Init();
            Evenhand();
        }
        #endregion

        #region 初始化
        /// <summary>
        /// 初始化
        /// </summary>
        void Init()
        {
            AcceptButton = bt_Conn_Test;
            MaximizeBox = false;
            MaximumSize = MinimumSize = Size;
        }
        #endregion

        #region 事件源绑定
        /// <summary>
        /// 事件源绑定
        /// </summary>
        void Evenhand()
        {
            //单选按钮
            rd_SqlServer.Click += new EventHandler(rd_Click);
            rd_Windows.Click += new EventHandler(rd_Click);
            //按钮
            bt_Conn_Test.Click += new EventHandler(bt_Click);
            bt_Login.Click += new EventHandler(bt_Click);
        }
        #endregion

        #region 控件事件

        #region 选中单选按钮
        /// <summary>
        /// 选中单选按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        #endregion

        #region 单击按钮
        /// <summary>
        /// 单击按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
                    if (rd_SqlServer.Checked)//sqlserver登陆
                    {
                        using (SqlConnection conn = new SqlConnection(string.Format(@"server={0};database=master;uid={1};pwd={2}",
                            txt_Server.Text.Trim(), txt_Uid.Text.Trim(), txt_Pwd.Text.Trim())))
                        {
                            try
                            { sqlconnstr = conn.ConnectionString; conn.Open(); bt_Login.Enabled = true; AcceptButton = bt_Login; }
                            catch
                            { MessageBox.Show("服务器连接失败!"); }
                        }
                    }
                    else//windows身份验证过
                    {
                        using (SqlConnection conn = new SqlConnection(string.Format(@"Data Source={0};database = master;Integrated security = true", txt_Server.Text.Trim())))
                        {
                            try
                            { sqlconnstr = conn.ConnectionString; conn.Open(); bt_Login.Enabled = true; AcceptButton = bt_Login; }
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
        #endregion

        private void bt_Conn_Test_Click(object sender, EventArgs e)
        {

        }

        #endregion
    }
}
