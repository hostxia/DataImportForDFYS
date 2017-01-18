using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using AfterVerificationCodeImport.Comm;
using AfterVerificationCodeImport.Demand;
using AfterVerificationCodeImport.Four;
using AfterVerificationCodeImport.Seven;
using AfterVerificationCodeImport.Nine;
using System.Data.OleDb;

namespace AfterVerificationCodeImport
{
    public partial class Frm_ReadExcel : Form
    { 
        private SqlConnection conn; 
        private int OutFileID;
        private int InFileID;
        private string connstr;

        public Frm_ReadExcel(string _connstr)
        {
            connstr = _connstr;
            InitializeComponent();
            Load += new EventHandler(Frm_ReadExcel_Load);
        }

        readonly DBHelper _dbHelper = new DBHelper();
        //读取发文和来文OID
        private void getType()
        {
             conn.Open();
            conn.ChangeDatabase(com_databasename.Text.Trim()); //重新指定数据库  
            OutFileID = _dbHelper.GetbySql("select OID FROM dbo.XPObjectType WHERE TypeName LIKE 'DataEntities.Element.Files.OutFile'",com_databasename.Text, conn);
            InFileID = _dbHelper.GetbySql("select OID FROM dbo.XPObjectType WHERE TypeName LIKE 'DataEntities.Element.Files.InFile'",com_databasename.Text, conn);
            conn.Close();
        }
        private void Frm_ReadExcel_Load(object sender, EventArgs e)
        {
            Init();
            EventHand();
        }
        private void Init()
        {
            conn = new SqlConnection(connstr); //SqlConnection实例化
            MaximizeBox = false; //禁用最小化
            MaximumSize = MinimumSize = Size; //固定当前大小
            txt_filepath.ReadOnly = true;
            //com_databasename.DropDownStyle = com_tablename.DropDownStyle = ComboBoxStyle.DropDownList; //下拉框只可选

            cbxSheet.DropDownStyle = cbxSheet.DropDownStyle = ComboBoxStyle.DropDownList; //下拉框只可选

            try
            {
                conn.Open();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"select name from sysdatabases";
                    SqlDataReader reader = cmd.ExecuteReader(); //获取所有数据库
                    while (reader.Read()) com_databasename.Items.Add(reader[0].ToString());
                    if (com_databasename.Items.Count > 0) com_databasename.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }
        }

        private void EventHand()
        {
            bt_see.Click += new EventHandler(bt_see_Click);
            //bt_ok.Click += new EventHandler(bt_ok_Click);
            FormClosing += new FormClosingEventHandler(Frm_ReadExcel_FormClosing);
        }

        #region 选择Excel文件

        /// <summary>
        /// 选择Excel文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bt_see_Click(object sender, EventArgs e)
        {
            try
            {
                var open = new OpenFileDialog();
                open.Filter = "所有文件(*.*)|*.*";
                open.ShowDialog(); //选择文件
                txt_filepath.Text = open.FileName;
                ReadExcel(open.FileName, 0);
                if (tablenames.Length > 0)
                {
                    cbxSheet.DataSource = tablenames;
                }
            }
            catch (Exception ex)
            {
                if (!ex.Message.Contains("无效的参数量"))
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        #endregion
        #region 窗体关闭

        /// <summary>
        /// 窗体关闭
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Frm_ReadExcel_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        #endregion

        #region 自定义方法
           private string[] tablenames;
        private DataSet ds;
        /// <summary>
        /// 读取Excel
        /// </summary>
        /// <param name="filename">文件名(含格式)</param>
        /// <param name="index">页数(默认0:第一页)</param>
        private void ReadExcel(string filename, int index)
        {
             var strConn = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + filename +
                             ";Extended Properties='Excel 12.0;'"; //此連接可以操作.xls與.xlsx文件
            using (var conn = new OleDbConnection(strConn))
            {
                conn.Open();
                DataTable table = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                 tablenames = new string[table.Rows.Count];
                for (int i = 0; i < table.Rows.Count; i++) tablenames[i] = table.Rows[i][2].ToString(); //获取Excel的表名
                if (tablenames.Length <= 0)
                {
                    MessageBox.Show("Excel中没有表!");
                    return;
                }
                using (OleDbCommand cmd = conn.CreateCommand())
                {       
                    lb_tablename.Text = "表名:" + tablenames[index].Substring(0, tablenames[index].Length - 1);
                    cmd.CommandText = "select * from [" + tablenames[index] + "]";
                    ds = new DataSet();
                    using (var da = new OleDbDataAdapter(cmd))
                    {
                        da.Fill(ds, tablenames[index]);
                        dgv_show.DataSource = ds.Tables[0];
                    }
                }
            }
        }

        #endregion

        //Sheet标签切换
        private void cbxSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReadExcel(txt_filepath.Text, cbxSheet.SelectedIndex);
        }

        
        #region  解释
        //s_IOType I：来文 O：发文 T：其它文件 
        //s_ClientGov  C: 客户  O: 官方 
        #endregion

        private int lostNum;
        private int sumCount;
        //导数据
        private void bt_ok_Click(object sender, EventArgs e)
        {
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("============="+xradioRegOnline.EditValue+" Start =================");
            getType();
            string mgr = string.Empty;
            sumCount = 0;
            lostNum = 0;
            var xfrmProcess = new ProBar();
            try
            {
                conn = new SqlConnection(connstr.Replace("master",com_databasename.Text));
                DateTime begin = DateTime.Now;
                //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("开始时间:" + begin);
                conn.Open();
                conn.ChangeDatabase(com_databasename.Text.Trim()); //重新指定数据库 
                var table = (DataTable)dgv_show.DataSource; //获取要Copy的数据源

                xfrmProcess.progressBarControl.Properties.Maximum = table.Rows.Count;
                xfrmProcess.progressBarControl.Properties.Minimum = 0;
                xfrmProcess.lbTotalSelected.Text = table.Rows.Count.ToString();
                xfrmProcess.Show();

                #region 客户基本信息
                if (xradioRegOnline.EditValue.ToString() == "客户信息")
                {
                    var _dealingClientData = new dealingClient();
                    string _selectValue = cbxSheet.SelectedValue.ToString();
                    if (_selectValue.Contains("联系人"))
                    {
                        DataRow[] newtableRow = table.Select(" [客户代码] is not null and [客户代码]<>''");
                        for (int i = 0; i < newtableRow.Length; i++)
                        {
                            sumCount++;
                            xfrmProcess.progressBarControl.Position = sumCount;
                            xfrmProcess.lbSuccess.Text = sumCount.ToString();
                            xfrmProcess.Refresh();
                            int resultNum = _dealingClientData.AddClientContact(newtableRow[i], i, com_databasename.Text, conn);//
                            if (resultNum == 0)
                            {
                                lostNum++;
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            sumCount++;
                            xfrmProcess.Show();
                            xfrmProcess.progressBarControl.Position = sumCount;
                            xfrmProcess.lbSuccess.Text = sumCount.ToString();
                            xfrmProcess.Refresh();

                            int resultNum = 0;
                            if (_selectValue.Contains("基本信息"))
                            {
                                resultNum = _dealingClientData.InsertTCstmrClient(table.Rows[i], i, com_databasename.Text, conn);
                            }
                            else if (_selectValue.Contains("地址"))
                            {
                                resultNum = _dealingClientData.AddClientAddress(table.Rows[i], i, com_databasename.Text, conn);
                            }
                            else if (_selectValue.Contains("财务信息"))
                            {
                                resultNum = _dealingClientData.AddBill(table.Rows[i],i,"",com_databasename.Text, conn);
                            }
                            if (resultNum == 0)
                            {
                                lostNum++;
                            }
                        }
                    }
                    _dealingClientData.updateApplicantandClient(com_databasename.Text, conn);
                }
                #endregion

                #region 申请人
                else if (xradioRegOnline.EditValue.ToString() == "申请人信息")
                {
                    var _dealingApplicant = new dealingApplicant();
                    string _selectValue = cbxSheet.SelectedValue.ToString();
                    if (_selectValue.Contains("联系人"))
                    {
                        DataRow[] newtableRow = table.Select(" [申请人代码] is not null and [申请人代码]<>''");
                        for (int i = 0; i < newtableRow.Length; i++)
                        {
                            sumCount++;
                            xfrmProcess.progressBarControl.Position = sumCount;
                            xfrmProcess.lbSuccess.Text = sumCount.ToString();
                            xfrmProcess.Refresh();
                            int resultNum = _dealingApplicant.AddAppContact(newtableRow[i], i, com_databasename.Text, conn);//
                            if (resultNum == 0)
                            {
                                lostNum++;
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            sumCount++;
                            xfrmProcess.progressBarControl.Position = sumCount;
                            xfrmProcess.lbSuccess.Text = sumCount.ToString(CultureInfo.InvariantCulture);
                            xfrmProcess.Refresh();

                            int resultNum = 0;
                            if (_selectValue.Contains("基本信息"))
                            {
                                resultNum = _dealingApplicant.InsertTCstmrApplicant(table.Rows[i], i, com_databasename.Text, conn);
                            }
                            else if (_selectValue.Contains("地址"))
                            {
                                resultNum = _dealingApplicant.AddAppAddress(table.Rows[i], i, com_databasename.Text, conn);
                            }
                            else if (_selectValue.Contains("财务信息"))
                            {
                                resultNum = _dealingApplicant.AddAppBill(table.Rows[i],i,"", com_databasename.Text, conn);
                            }
                            if (resultNum == 0)
                            {
                                lostNum++;
                            }
                        }
                    }
                    var _dealingClientData = new dealingClient();
                    _dealingClientData.updateApplicantandClient(com_databasename.Text, conn);
                }
                #endregion

                #region 总委托书号
                else if (xradioRegOnline.EditValue.ToString() == "总委托书号")
                {
                    var _dealingApplicant = new dealingApplicant();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        
                        int resultNum = _dealingApplicant.UpdateTotalCommissionNumber(table.Rows[i], i,com_databasename.Text, conn);
                        if (resultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region  外代理信息
                else if (xradioRegOnline.EditValue.ToString() == "外代理信息")
                {
                    var _dealingAgency = new dealingAgency();
                    string _selectValue = cbxSheet.SelectedValue.ToString();
                    if (_selectValue.Contains("联系人"))
                    {
                        DataRow[] newtableRow = table.Select(" [外代理] is not null  and [外代理]<>'' ");
                        for (int i = 0; i < newtableRow.Length; i++)
                        {
                            sumCount++;
                            xfrmProcess.progressBarControl.Position = sumCount;
                            xfrmProcess.lbSuccess.Text = sumCount.ToString();
                            xfrmProcess.Refresh();
                            int resultNum = _dealingAgency.AddAgencyContact(newtableRow[i], i,com_databasename.Text,conn);//
                            if (resultNum == 0)
                            {
                                lostNum++;
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            sumCount++;
                            xfrmProcess.progressBarControl.Position = sumCount;
                            xfrmProcess.lbSuccess.Text = sumCount.ToString();
                            xfrmProcess.Refresh();

                            int resultNum = 0;
                            if (_selectValue.Contains("基本信息"))
                            {
                                resultNum = _dealingAgency.InsertTCstmrCoopAgency(table.Rows[i], i,com_databasename.Text, conn);
                            }
                            else if (_selectValue.Contains("地址"))
                            {
                                resultNum = _dealingAgency.AddAgencyAddress(table.Rows[i], i,com_databasename.Text, conn);
                            }
                            else if (_selectValue.Contains("财务信息"))
                            {
                                resultNum = _dealingAgency.AddAgencyBill(table.Rows[i],i,com_databasename.Text,conn);
                            }
                            if (resultNum == 0)
                            {
                                lostNum++;
                            }
                        }
                    }
                }
                #endregion

                #region  实体信息表

                else if (xradioRegOnline.EditValue.ToString() == "实体信息表")
                {
                    var _dealingthreeData = new dealingthreeData();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int ResultNum = _dealingthreeData.TCaseApplicant(i, table.Rows[i],com_databasename.Text, conn);
                        if (ResultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 更新申请人译名
                else if (xradioRegOnline.EditValue.ToString() == "更新申请人译名")
                {
                    var _dealingthreeData = new dealingthreeData();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = _dealingthreeData.UpdateApplicant(table.Rows[i], i, com_databasename.Text, conn);
                        if (resultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 国内-专利数据补充导入

                else if (xradioRegOnline.EditValue.ToString() == "国内-专利数据补充导入")
                {
                    var _dealingCasePantent = new dealingCasePantent();
                    #region 国内-专利数据补充导入

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int ResultNum = _dealingCasePantent.InsertPatented(i, table.Rows[i], com_databasename.Text, conn);
                        if (ResultNum == 0)
                        {
                            lostNum++;
                        }
                    }

                    #endregion
                }
                #endregion

                #region 国内-OA数据补充导入
                else if (xradioRegOnline.EditValue.ToString() == "国内-OA数据补充导入")
                {
                    var _dealingOAData = new dealingOAData();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int resultNum = _dealingOAData.OAData(i, table.Rows[i], OutFileID, InFileID, com_databasename.Text, conn);
                        if (resultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region  国内-OA数据补充导入-案件处理人
                else if (xradioRegOnline.EditValue.ToString() == "案件处理人")
                {
                    var _dealingOAData = new dealingOAData();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString(CultureInfo.InvariantCulture);
                        xfrmProcess.Refresh();
                        int result = _dealingOAData.InsertCaseAttorney(table, i, table.Rows[i], com_databasename.Text, conn);
                        if (result == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 国内-优先权
                else if (xradioRegOnline.EditValue.ToString() == "优先权")
                {
                    var _dealingCasePriority = new dealingCasePriority();
                    DataRow[] NewTable = table.Select("", "优先权号  ");
                    for (int i = 0; i < NewTable.Length; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int ReturnNum = _dealingCasePriority.TPCasePriority(i, NewTable[i], com_databasename.Text, conn);
                        if (ReturnNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region  香港优先权
                else if (xradioRegOnline.EditValue.ToString() == "优先权（香港）")
                {
                    var _dealingCasePriority = new dealingCasePriority();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = _dealingCasePriority.HongKangPriority(i, table.Rows[i], com_databasename.Text, conn);
                        if (resultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region  国外-法律信息及日志表
                else if (xradioRegOnline.EditValue.ToString() == "国外-法律信息及日志表")
                {
                    var _dealingTCaseLaw = new dealingTCaseLaw();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int ResultNum = _dealingTCaseLaw.TCaseLaw(i, table.Rows[i], InFileID, OutFileID, com_databasename.Text, conn);
                        if (ResultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region  国外-国外库发明人表
                else if (xradioRegOnline.EditValue.ToString() == "国外-国外库发明人表")
                {
                    var _dealingCaseInventor = new dealingCaseInventor();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int ResultNum = _dealingCaseInventor.TPCaseInventor(i, table.Rows[i], com_databasename.Text, conn);
                        if (ResultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 香港-专利数据补充导入
                else if (xradioRegOnline.EditValue.ToString() == "香港-专利数据补充导入")
                {
                    var _dealingCasePantent = new dealingCasePantent();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int Result = _dealingCasePantent.HongKang(i, table.Rows[i], OutFileID,com_databasename.Text, conn);
                        if (Result == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region  根据业务类型修改申请方式
                else if (xradioRegOnline.EditValue.ToString() == "申请方式")
                {
                    var _dealingTCodeBusinessType = new dealingTCodeBusinessType();
                    string _selectValue = cbxSheet.SelectedValue.ToString();
                    if (_selectValue.Contains("美国"))
                    {
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            sumCount++;
                            xfrmProcess.progressBarControl.Position = sumCount;
                            xfrmProcess.lbSuccess.Text = sumCount.ToString();
                            xfrmProcess.Refresh();
                            if (table.Rows[i]["申请国家"].Equals("美国"))
                            {
                                int resultNum = _dealingTCodeBusinessType.UpdateUSAType(table.Rows[i], i, com_databasename.Text, conn);
                                if (resultNum == 0)
                                {
                                    lostNum++;
                                }
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            sumCount++;
                            xfrmProcess.progressBarControl.Position = sumCount;
                            xfrmProcess.lbSuccess.Text = sumCount.ToString();
                            xfrmProcess.Refresh();

                            int resultNum = _dealingTCodeBusinessType.UpdateType(table.Rows[i], i, com_databasename.Text, conn);
                            if (resultNum == 0)
                            {
                                lostNum++;
                            }
                        }
                    }
                }
                #endregion

                #region 案件部门导入
                else if (xradioRegOnline.EditValue.ToString() == "部门核对")
                {
                    var _dealingDepartment = new dealingDepartment();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = _dealingDepartment.UpdateOrg(table.Rows[i], i, com_databasename.Text, conn);
                        if (resultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 国外-国外库时限备注表
                else if (xradioRegOnline.EditValue.ToString() == "国外-国外库时限备注表")
                {
                    var _dealingCaseMemo = new dealingCaseMemo();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int RsultNum = _dealingCaseMemo.Case_Memo(i, table.Rows[i], OutFileID,com_databasename.Text, conn);
                        if (RsultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 国内-收文数据导入
                if (xradioRegOnline.EditValue.ToString() == "国内-收文数据导入")
                {
                    var _tMainFiles = new T_MainFiles();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int ResultNum = _tMainFiles.InsertFileIn(i, table.Rows[i], InFileID, com_databasename.Text, conn);
                        if (ResultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 年费
                else if (xradioRegOnline.EditValue.ToString() == "年费")
                {
                    var _dealingFee = new dealingFee();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int resultNum = _dealingFee.InsertFee(i, table.Rows[i], com_databasename.Text, conn);
                        if (resultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 任务时限
                else if (xradioRegOnline.EditValue.ToString() == "任务时限")
                {
                    var _dealingTask = new dealingTask();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int ResultNum = _dealingTask.ImportTask(i, table.Rows[i],com_databasename.Text, conn);
                        if (ResultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 分案信息
                else if (xradioRegOnline.EditValue.ToString() == "分案信息")
                {
                    var _dealingCasePantent = new dealingCasePantent();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = _dealingCasePantent.UpdateCaseDivisionInfo(table.Rows[i], i, com_databasename.Text, conn);
                        if (resultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 相关案件-双申
                else if (xradioRegOnline.EditValue.ToString() == "相关案件-双申")
                {
                    var _dealingCaseToCase = new dealingCaseToCase();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString(CultureInfo.InvariantCulture);
                        xfrmProcess.Refresh();
                        int result = _dealingCaseToCase.InsertDoubleShen(table.Rows[i], i, com_databasename.Text, conn);
                        if (result == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 自定义属性
                else if (xradioRegOnline.EditValue.ToString() == "自定义属性")
                {
                    var _dealingCodeCaseCustomField = new dealingCodeCaseCustomField();
                    string _selectValue = cbxSheet.SelectedValue.ToString();
                    if (_selectValue.Contains("专利数据补充"))
                    {
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            sumCount++;
                            xfrmProcess.progressBarControl.Position = sumCount;
                            xfrmProcess.lbSuccess.Text = sumCount.ToString();
                            xfrmProcess.Refresh();
                            int resultNum = _dealingCodeCaseCustomField.InsertPantentCodeCaseCustomField("权项数", table.Rows[i]["权项数"].ToString(), table.Rows[i]["我方文号"].ToString(), i, com_databasename.Text, conn);
                            resultNum = _dealingCodeCaseCustomField.InsertPantentCodeCaseCustomField("总字数", table.Rows[i]["总字数"].ToString(), table.Rows[i]["我方文号"].ToString(), i, com_databasename.Text, conn);
                            if (resultNum == 0)
                            {
                                lostNum++;
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            sumCount++;
                            xfrmProcess.progressBarControl.Position = sumCount;
                            xfrmProcess.lbSuccess.Text = sumCount.ToString();
                            xfrmProcess.Refresh();
                            int resultNum = _dealingCodeCaseCustomField.InsertCodeCaseCustomField(table.Rows[i], i, com_databasename.Text, conn);
                            if (resultNum == 0)
                            {
                                lostNum++;
                            }
                        }
                    }
                }
                #endregion

                #region 备注
                else if (xradioRegOnline.EditValue.ToString() == "备注")
                {
                    var _dealingCaseMemo = new dealingCaseMemo();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = _dealingCaseMemo.InsertTCaseMemo(table.Rows[i], i, com_databasename.Text, conn);
                        if (resultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 要求/发文
                else if (xradioRegOnline.EditValue.ToString() == "要求/发文")
                {
                    var _dealingIPS = new dealingIPS();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = _dealingIPS.sGetType(table.Rows[i], i, OutFileID, com_databasename.Text, conn);
                        if (resultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 档案位置
                else if (xradioRegOnline.EditValue.ToString() == "档案位置")
                {
                    var _dealingIPS = new dealingIPS();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = _dealingIPS.InsertFileLocation(table.Rows[i], i, com_databasename.Text, conn);
                        if (resultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 澳门案件
                else if (xradioRegOnline.EditValue.ToString() == "澳门案件-澳门延伸")
                {
                    var _dealingIPS = new dealingIPS();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString(CultureInfo.InvariantCulture);
                        xfrmProcess.Refresh();
                        int result = _dealingIPS.InsertMacaoApplication(table.Rows[i], i, com_databasename.Text, conn);
                        if (result == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 集团客户代码-相关客户
                else if (xradioRegOnline.EditValue.ToString() == "相关客户")
                {
                    DataRow[] newtableRow = table.Select(" [客户代码] is not null");
                    var _dealingRelatedcustomers = new dealingRelatedcustomers();
                    for (int i = 0; i < newtableRow.Length; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = _dealingRelatedcustomers.UpdateRelatedcustomers(newtableRow[i], i, com_databasename.Text, conn);
                        if (resultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                else if (xradioRegOnline.EditValue.ToString() == "相关客户要求")
                {
                    var _dealingRelatedcustomers = new dealingRelatedcustomers();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = _dealingRelatedcustomers.InsertClientDemnd(table.Rows[i], i, com_databasename.Text, conn);
                        if (resultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region 要求配置
                else if (xradioRegOnline.EditValue.ToString() == "要求配置")
                {
                    var _dealingClientandCase = new dealingClientandCase();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = _dealingClientandCase.InsertDemand(table.Rows[i], i,com_databasename.Text, conn);//增加客户、申请人、客户-申请人要求
                        _dealingClientandCase.InsertCaseDemand(table.Rows[i], com_databasename.Text, conn); //拷贝案子要求
                        if (resultNum == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region  客户要求
                else if (xradioRegOnline.EditValue.ToString() == "客户要求")
                {
                    var _dealingClientDemand = new dealingClientDemand();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString(CultureInfo.InvariantCulture);
                        xfrmProcess.Refresh();
                        int result = _dealingClientDemand.InsertDemandClient(table.Rows[i], i, com_databasename.Text, conn);
                        if (result == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                #region  Update
                else if (xradioRegOnline.EditValue.ToString() == "Update")
                {
                    var _dealingCasePantent = new dealingCasePantent();
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString(CultureInfo.InvariantCulture);
                        xfrmProcess.Refresh();
                        int result = _dealingCasePantent.UpdateCase(table.Rows[i], i, com_databasename.Text, conn);
                        if (result == 0)
                        {
                            lostNum++;
                        }
                    }
                }
                #endregion

                else //整体不需要转换导入
                {
                    //SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                    //foreach (SqlBulkCopyColumnMapping columnsmapping in sqlBulkCopyparameters)
                    //    bulkcopy.ColumnMappings.Add(columnsmapping); //加载文件与表的列名 
                    //bulkcopy.DestinationTableName = com_tablename.Text.Trim(); //指定目标表的表名
                    //bulkcopy.WriteToServer(table); //将table Copy到数据库 
                }
                DateTime end = DateTime.Now;
                TimeSpan ts = end.Subtract(begin).Duration();
                xfrmProcess.progressBarControl.Position = xfrmProcess.progressBarControl.Properties.Maximum;
                xfrmProcess.lbTotalTime.Text = ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒";
                xfrmProcess.lbTotalSum.Text = lostNum + "条";

                Application.DoEvents();
                Thread.Sleep(8000);
                xfrmProcess.Focus();
                xfrmProcess.Close();
                mgr = string.Format("服务器:{0}\n数据库:{1}\n共复制{2}条数据\n{3}条数据未找到“我方卷号”\n耗时{4}天{5}小时{6}分{7}秒",
                                    conn.DataSource, com_databasename.Text, table.Rows.Count, lostNum, ts.Days, ts.Hours,
                                    ts.Minutes, ts.Seconds);
                conn.Close();

                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(mgr);
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============" + xradioRegOnline.EditValue + " End =================");

            }
            catch (Exception ex)
            {
                Application.DoEvents();
                Thread.Sleep(3000);
                xfrmProcess.Focus();
                xfrmProcess.Close();
                conn.Close();
                MessageBox.Show(ex.Message);
            }
        }

        //相关客户要求
        private void button6_Click(object sender, EventArgs e)
        {
            var xfrmProcess = new ProBar();

            DateTime begin = DateTime.Now;
            conn.Open();
            conn.ChangeDatabase(com_databasename.Text.Trim()); //重新指定数据库 

            var _dealingCaseDemand = new dealingCaseDemand();

            DataTable table = _dbHelper.GetDataTablebySql("SELECT n_CaseID FROM dbo.TCase_Base", conn);

            xfrmProcess.progressBarControl.Properties.Maximum = table.Rows.Count;
            xfrmProcess.progressBarControl.Properties.Minimum = 0;
            xfrmProcess.lbTotalSelected.Text = table.Rows.Count.ToString();
            xfrmProcess.Show();
            for (var i = 0; i < table.Rows.Count; i++)
            {
                sumCount++;
                xfrmProcess.progressBarControl.Position = sumCount;
                xfrmProcess.lbSuccess.Text = sumCount.ToString();
                xfrmProcess.Refresh();
                if (_dealingCaseDemand.InsertDemand(table.Rows[i]["n_CaseID"].ToString(), i, com_databasename.Text, conn) <= 0)
                {
                    lostNum++;
                }
            }
            DateTime end = DateTime.Now;
            TimeSpan ts = end.Subtract(begin).Duration();
            xfrmProcess.progressBarControl.Position = xfrmProcess.progressBarControl.Properties.Maximum;
            xfrmProcess.lbTotalTime.Text = ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒";
            xfrmProcess.lbTotalSum.Text = lostNum + "条";

            Application.DoEvents();
            Thread.Sleep(8000);
            xfrmProcess.Focus();
            xfrmProcess.Close();
            string mgr = string.Format("服务器:{0}\n数据库:{1}\n共复制{2}条数据\n{3}条数据未找到“我方卷号”\n耗时{4}天{5}小时{6}分{7}秒",
                                conn.DataSource, com_databasename.Text, table.Rows.Count, lostNum, ts.Days, ts.Hours,
                                ts.Minutes, ts.Seconds);
            conn.Close();
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(mgr);
        }

        #region 按钮事件
        //更新数据
        private void button2_Click(object sender, EventArgs e)
        {
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============更新数据按钮 Start =================");
            sumCount = 0;
            conn.Open();
            conn.ChangeDatabase(com_databasename.Text.Trim()); //重新指定数据库 
            DateTime begin = DateTime.Now;

            #region
            string strSql = " update TCstmr_Applicant  set n_Country=(select n_Country from TCstmr_Client where TCstmr_Client.n_ClientID=TCstmr_Applicant.n_ClientID)";
            _dbHelper.InsertbySql(strSql, 0, com_databasename.Text, conn);

            strSql = "SELECT n_CaseID FROM dbo.TCase_Base";
            DataTable table = _dbHelper.GetDataTablebySql(strSql, conn);

            var xfrmProcess = new ProBar();
            xfrmProcess.progressBarControl.Properties.Maximum = table.Rows.Count;
            xfrmProcess.progressBarControl.Properties.Minimum = 0;
            xfrmProcess.lbTotalSelected.Text = table.Rows.Count.ToString();
            xfrmProcess.Show();

            for (int i = 0; i < table.Rows.Count; i++)
            {
                sumCount++;
                xfrmProcess.progressBarControl.Position = sumCount;
                xfrmProcess.lbSuccess.Text = sumCount.ToString();
                xfrmProcess.Refresh();
                Application.DoEvents();
                strSql = string.Format(@"
SELECT dbo.TCase_Applicant.n_CaseID,dbo.TCstmr_Applicant.n_Country,n_Sequence
FROM dbo.TCase_Applicant 
LEFT JOIN dbo.TCstmr_Applicant ON dbo.TCstmr_Applicant.n_AppID = dbo.TCase_Applicant.n_ApplicantID
WHERE ISNULL(dbo.TCstmr_Applicant.n_Country,0) > 0 AND n_CaseID = '{0}'
ORDER BY n_Sequence ASC", table.Rows[i]["n_CaseID"]);
                DataTable tableApplicant = _dbHelper.GetDataTablebySql(strSql, conn);
                if (tableApplicant.Rows.Count > 0)
                {
                    string nCountry = tableApplicant.Rows[0]["n_Country"].ToString();
                    strSql = " UPDATE TCase_Base SET n_AppCountry=" + nCountry + " WHERE n_CaseID=" + table.Rows[i]["n_CaseID"];
                    _dbHelper.InsertbySql(strSql, i, com_databasename.Text, conn);
                }
            }
            #endregion

            //更新案件流向
            strSql = @"
UPDATE dbo.TCase_Base SET s_FlowDirection = CASE 
WHEN AppCountry.s_CountryCode = 'CN' AND RegCountry.s_CountryCode = 'CN' THEN 'II'
WHEN AppCountry.s_CountryCode = 'CN' AND RegCountry.s_CountryCode <> 'CN' THEN 'IO'
WHEN AppCountry.s_CountryCode <> 'CN' AND RegCountry.s_CountryCode = 'CN' THEN 'OI'
WHEN AppCountry.s_CountryCode <> 'CN' AND RegCountry.s_CountryCode <> 'CN' THEN 'OO'
WHEN ISNULL(AppCountry.s_CountryCode,'') = '' OR ISNULL(RegCountry.s_CountryCode,'') = '' THEN ''
ELSE '' END 
FROM dbo.TCase_Base AS BaseCase
LEFT JOIN dbo.TCode_Country AS AppCountry ON AppCountry.n_ID = BaseCase.n_AppCountry
LEFT JOIN dbo.TCode_Country AS RegCountry ON RegCountry.n_ID = BaseCase.n_RegCountry
WHERE n_CaseID = BaseCase.n_CaseID 
";
            _dbHelper.InsertbySql(strSql, 0, com_databasename.Text, conn);

            //根据是否填写了申请号更新案件提交状态
            strSql = @"
UPDATE dbo.TCase_Base SET s_SubmitStatus = 'Y'
FROM dbo.TPCase_Patent
INNER JOIN dbo.TPCase_LawInfo ON dbo.TPCase_LawInfo.n_ID = dbo.TPCase_Patent.n_LawID
WHERE dbo.TCase_Base.n_CaseID = dbo.TPCase_Patent.n_CaseID AND ISNULL(dbo.TPCase_LawInfo.s_AppNo,'') <> ''
";
            _dbHelper.InsertbySql(strSql, 0, com_databasename.Text, conn);

            strSql = @"
UPDATE dbo.TCase_Base SET s_SubmitStatus = 'N'
FROM dbo.TPCase_Patent
INNER JOIN dbo.TPCase_LawInfo ON dbo.TPCase_LawInfo.n_ID = dbo.TPCase_Patent.n_LawID
WHERE dbo.TCase_Base.n_CaseID = dbo.TPCase_Patent.n_CaseID AND ISNULL(dbo.TPCase_LawInfo.s_AppNo,'') = ''
";
            _dbHelper.InsertbySql(strSql, 0, com_databasename.Text, conn);

            //根据注册国家更新案件递交机构
            strSql = @"
UPDATE dbo.TCase_Base SET n_OfficeID = dbo.TCode_Official.n_ID
FROM dbo.TCode_Official
WHERE dbo.TCase_Base.n_RegCountry = dbo.TCode_Official.n_Country AND dbo.TCode_Official.s_IPType LIKE '%P%'
";
            _dbHelper.InsertbySql(strSql, 0, com_databasename.Text, conn);

            //PCT国际案件：PCT国际申请没有填写PCT申请号，只填写了申请号
            strSql = "update TPCase_LawInfo set s_PCTAppNo=s_AppNo    where n_ID in (select n_LawID from TPCase_Patent where n_CaseID in (select n_CaseID from tcase_base where n_BusinessTypeID in (select n_ID from tcode_BusinessType where s_Name='PCT国际阶段申请')))  and s_AppNo  is not null  and isnull(s_PCTAppNo,'')=''";
            _dbHelper.InsertbySql(strSql, 0, com_databasename.Text, conn);

            //通过原案我方文号从法律信息中获取数据更新到原案信息中
            strSql = @"
UPDATE dbo.TPCase_OrigPatInfo SET 
dt_AppDate = dbo.TPCase_LawInfo.dt_AppDate,
s_AppNo = dbo.TPCase_LawInfo.s_AppNo,
n_ClaimCount = dbo.TPCase_LawInfo.n_ClaimCount,
s_PubNo = dbo.TPCase_LawInfo.s_PubNo,
dt_PubDate = dbo.TPCase_LawInfo.dt_PubDate,
s_PubVolume = dbo.TPCase_LawInfo.s_PubVolume,
s_PubGazette = dbo.TPCase_LawInfo.s_PubGazette,
s_PubClass1 = dbo.TPCase_LawInfo.s_PubClass1,
s_PubClass2 = dbo.TPCase_LawInfo.s_PubClass2,
dt_SubstantiveExamDate = dbo.TPCase_LawInfo.dt_SubstantiveExamDate,
s_IssuedPubNo = dbo.TPCase_LawInfo.s_IssuedPubNo,
dt_IssuedPubDate = dbo.TPCase_LawInfo.dt_IssuedPubDate,
s_IssuedPubVolume = dbo.TPCase_LawInfo.s_IssuedPubVolume,
s_IssuedPubGazette = dbo.TPCase_LawInfo.s_IssuedPubGazette,
s_IssuedPubClass1 = dbo.TPCase_LawInfo.s_IssuedPubClass1,
s_IssuedPubClass2 = dbo.TPCase_LawInfo.s_IssuedPubClass2,
s_CertfNo = dbo.TPCase_LawInfo.s_CertfNo,
dt_CertfDate = dbo.TPCase_LawInfo.dt_CertfDate,
s_PCTAppNo = dbo.TPCase_LawInfo.s_PCTAppNo,
dt_PctAppDate = dbo.TPCase_LawInfo.dt_PctAppDate,
s_PCTPubNo = dbo.TPCase_LawInfo.s_PCTPubNo,
dt_PctPubDate = dbo.TPCase_LawInfo.dt_PctPubDate,
dt_PctInNationDate = dbo.TPCase_LawInfo.dt_PctInNationDate,
n_PCTPubLan = dbo.TPCase_LawInfo.n_PCTPubLan,
dt_FirstCheckDate = dbo.TPCase_LawInfo.dt_FirstCheckDate,
s_PatentStatus = dbo.TPCase_Patent.s_PatentStatus,
s_Note = dbo.TPCase_Patent.s_Notes,
s_PatentName = dbo.TCase_Base.s_CaseName,
n_OrigRegCountry = dbo.TCase_Base.n_RegCountry
FROM dbo.TPCase_LawInfo
INNER JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_LawID = dbo.TPCase_LawInfo.n_ID
INNER JOIN dbo.TCase_Base ON dbo.TCase_Base.n_CaseID = dbo.TPCase_Patent.n_CaseID
WHERE dbo.TCase_Base.s_CaseSerial = dbo.TPCase_OrigPatInfo.s_CaseSerial
";
            _dbHelper.InsertbySql(strSql, 0, com_databasename.Text, conn);

            //从本案信息中查找相关信息更新到澳门案的个案信息中
            strSql = @"
UPDATE dbo.TPCase_MacaoApplication SET 
s_ParentCaseSerial = TPCase_OrigPatInfo.s_CaseSerial,
s_ParentCaseAppNo = TPCase_OrigPatInfo.s_AppNo,
s_ParentCaseCountry = (SELECT TOP 1 s_Name FROM dbo.TCode_Country WHERE n_ID = dbo.TPCase_OrigPatInfo.n_OrigRegCountry),
dt_AppDate = TPCase_LawInfo.dt_AppDate,
dt_GrantDate = TPCase_LawInfo.dt_IssuedPubDate
FROM dbo.TCase_Base
LEFT JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_CaseID = dbo.TCase_Base.n_CaseID
LEFT JOIN dbo.TPCase_LawInfo ON dbo.TPCase_LawInfo.n_ID = dbo.TPCase_Patent.n_LawID
LEFT JOIN dbo.TPCase_OrigPatInfo ON dbo.TPCase_OrigPatInfo.s_CaseSerial = dbo.TCase_Base.s_CaseSerial
WHERE dbo.TPCase_MacaoApplication.n_CaseID = dbo.TCase_Base.n_CaseID
";
            _dbHelper.InsertbySql(strSql, 0, com_databasename.Text, conn);

            DateTime end = DateTime.Now;
            TimeSpan ts = end.Subtract(begin).Duration();
            xfrmProcess.progressBarControl.Position = xfrmProcess.progressBarControl.Properties.Maximum;
            xfrmProcess.lbTotalTime.Text = ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒";
            xfrmProcess.lbTotalSum.Text = lostNum + "条";
            Application.DoEvents();
            Thread.Sleep(8000);
            xfrmProcess.Focus();
            xfrmProcess.Close();
            conn.Close();
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============" + lostNum + "条" + ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒" + "更新数据按钮 End =================");
        }

        #region 2.同步PCT案件关系
        private void btnPCT_Click(object sender, EventArgs e)
        {
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============同步PCT案件关系 Start =================");
            sumCount = 0;
            conn.Open();
            conn.ChangeDatabase(com_databasename.Text.Trim()); //重新指定数据库 
            DateTime begin = DateTime.Now;
            //1、查找业务类型为国际阶段申请案类型的案件
            DataTable PctTable =
                _dbHelper.GetDataTablebySql(
@"SELECT dbo.TCase_Base.n_CaseID,s_CaseSerial,dbo.TPCase_LawInfo.s_AppNo  
FROM dbo.TCase_Base
INNER JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_CaseID = dbo.TCase_Base.n_CaseID
INNER JOIN dbo.TPCase_LawInfo ON dbo.TPCase_LawInfo.n_ID = dbo.TPCase_Patent.n_LawID
WHERE n_BusinessTypeID IN (SELECT n_ID FROM dbo.TCode_BusinessType WHERE s_Name='PCT国际阶段申请') AND ISNULL(dbo.TPCase_LawInfo.s_AppNo,'') <> ''
UNION 
SELECT dbo.TCase_Base.n_CaseID,s_CaseSerial,dbo.TPCase_LawInfo.s_PCTAppNo  
FROM dbo.TCase_Base
INNER JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_CaseID = dbo.TCase_Base.n_CaseID
INNER JOIN dbo.TPCase_LawInfo ON dbo.TPCase_LawInfo.n_ID = dbo.TPCase_Patent.n_LawID
WHERE n_BusinessTypeID IN (SELECT n_ID FROM dbo.TCode_BusinessType WHERE s_Name='PCT国际阶段申请') AND ISNULL(dbo.TPCase_LawInfo.s_PCTAppNo,'') <> ''", conn);
            var xfrmProcess = new ProBar();
            xfrmProcess.progressBarControl.Properties.Maximum = PctTable.Rows.Count;
            xfrmProcess.progressBarControl.Properties.Minimum = 0;
            xfrmProcess.lbTotalSelected.Text = PctTable.Rows.Count.ToString();
            xfrmProcess.Show();

            if (PctTable.Rows.Count > 0)
            {
                //2、取每一条PCT国际申请案的申请号，在TPCase_LawInfo中查询s_PCTAppNo=PCT国际申请案的申请号的数据，取得案件ID
                for (int i = 0; i < PctTable.Rows.Count; i++)
                {
                    sumCount++;
                    xfrmProcess.progressBarControl.Position = sumCount;
                    xfrmProcess.lbSuccess.Text = sumCount.ToString();
                    xfrmProcess.Refresh();
                    Application.DoEvents();
                    int nCaseID = Convert.ToInt32(PctTable.Rows[i]["n_CaseID"].ToString());
                    string sAppNo = PctTable.Rows[i]["s_AppNo"].ToString();

                    if (!string.IsNullOrEmpty(sAppNo))
                    {
                        DataTable tablePCTN =
                            _dbHelper.GetDataTablebySql(
                                string.Format(
@"SELECT DISTINCT dbo.TCase_Base.n_CaseID
FROM dbo.TCase_Base
INNER JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_CaseID = dbo.TCase_Base.n_CaseID
INNER JOIN dbo.TPCase_LawInfo ON dbo.TPCase_LawInfo.n_ID = dbo.TPCase_Patent.n_LawID
WHERE n_BusinessTypeID NOT IN (SELECT n_ID FROM dbo.TCode_BusinessType WHERE s_Name='PCT国际阶段申请') 
AND dbo.TPCase_LawInfo.s_PCTAppNo = '{0}'", sAppNo), conn);

                        var listPCTNCaseID = tablePCTN.Rows.Cast<DataRow>().Select(dr => dr["n_CaseID"].ToString()).ToList();
                        var listPCTNCaseIDCopy = tablePCTN.Rows.Cast<DataRow>().Select(dr => dr["n_CaseID"].ToString()).ToList();
                        const string strSql = "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='PCT国家' AND s_MasterName='同PCT案' AND s_SlaveName='同PCT案' AND s_IPType='P'";
                        int n_ID = _dbHelper.GetbySql(strSql, com_databasename.Text, conn);
                        const string strSqlPct = "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='PCT' AND s_MasterName='PCT国际案' AND s_SlaveName='PCT国家案' AND s_IPType='P'";
                        int nIDPCT = _dbHelper.GetbySql(strSqlPct, com_databasename.Text, conn);

                        foreach (var exceptCasePCTNID in listPCTNCaseID)
                        {
                            foreach (var casePCTNID in listPCTNCaseIDCopy)
                            {
                                if (exceptCasePCTNID == casePCTNID) continue;
                                InsertIntoLaw(Convert.ToInt32(casePCTNID), Convert.ToInt32(exceptCasePCTNID), n_ID, i, "");
                            }
                            InsertIntoLaw(Convert.ToInt32(exceptCasePCTNID), nCaseID, nIDPCT, i, "");
                            listPCTNCaseIDCopy.Remove(exceptCasePCTNID);
                        }
                    }
                    else
                    {
                        Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("申请号为空：" +
                                                                                   PctTable.Rows[i]["s_CaseSerial"].
                                                                                       ToString());
                    }
                }
            }
            else
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("未查到PCT国际阶段申请");
            }
            DateTime end = DateTime.Now;
            TimeSpan ts = end.Subtract(begin).Duration();
            xfrmProcess.progressBarControl.Position = xfrmProcess.progressBarControl.Properties.Maximum;
            xfrmProcess.lbTotalTime.Text = ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒";
            xfrmProcess.lbTotalSum.Text = lostNum + "条";
            Application.DoEvents();
            Thread.Sleep(8000);
            xfrmProcess.Focus();
            xfrmProcess.Close();
            conn.Close();
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============" + lostNum + "条" + ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒" + "=================");
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============同步PCT案件关系 End =================");
        }

        //增加案件关系表
        private void InsertIntoLaw(int nCaseID, int HKNum, int n_ID, int i, string type)
        {
            int NUMS =
                _dbHelper.GetbySql("SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where n_CaseIDA=" + HKNum +
                           " and n_CaseIDB=" + nCaseID + " and n_CodeRelativeID=" + n_ID,com_databasename.Text, conn);
            if (!string.IsNullOrEmpty(type)) //同族
            {
                NUMS =
                    _dbHelper.GetbySql("SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where ((n_CaseIDA=" + HKNum +
                               " and n_CaseIDB=" + nCaseID + ") or (n_CaseIDA=" + nCaseID + " and n_CaseIDB=" + HKNum +
                               ") )and n_CodeRelativeID=" + n_ID, com_databasename.Text, conn);
            }
            var _tMainFiles = new T_MainFiles();
            if (NUMS <= 0 && HKNum > 0)
            {
                _tMainFiles.InsertTCaseCaseRelative(HKNum, nCaseID, n_ID, 0, i, com_databasename.Text, conn);
            }
        }
        #endregion

        #region 3.IDS案件关系
        private void IDSCaseRelative()
        {
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============IDS案件关系 Start =================");
            sumCount = 0;
            conn.Open();
            conn.ChangeDatabase(com_databasename.Text.Trim()); //重新指定数据库 
            DateTime begin = DateTime.Now;

            #region
            //1、从系统中筛选出所有注册国家是美国的案件。
            string strSql = @"
SELECT dbo.TCase_Base.n_CaseID,s_CaseSerial,dbo.TPCase_LawInfo.s_AppNo  
FROM dbo.TCase_Base
INNER JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_CaseID = dbo.TCase_Base.n_CaseID
INNER JOIN dbo.TPCase_LawInfo ON dbo.TPCase_LawInfo.n_ID = dbo.TPCase_Patent.n_LawID
WHERE n_RegCountry IN (SELECT n_ID FROM TCode_Country WHERE S_Name='美国')
UNION ALL
SELECT dbo.TCase_Base.n_CaseID,s_CaseSerial,dbo.TPCase_LawInfo.s_PCTAppNo  
FROM dbo.TCase_Base
INNER JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_CaseID = dbo.TCase_Base.n_CaseID
INNER JOIN dbo.TPCase_LawInfo ON dbo.TPCase_LawInfo.n_ID = dbo.TPCase_Patent.n_LawID
WHERE n_RegCountry IN (SELECT n_ID FROM TCode_Country WHERE S_Name='美国')";
            DataTable table = _dbHelper.GetDataTablebySql(strSql, conn);

            var xfrmProcess = new ProBar();
            xfrmProcess.progressBarControl.Properties.Maximum = table.Rows.Count;
            xfrmProcess.progressBarControl.Properties.Minimum = 0;
            xfrmProcess.lbTotalSelected.Text = table.Rows.Count.ToString();
            xfrmProcess.Show();

            for (int i = 0; i < table.Rows.Count; i++)
            {
                //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============总行数：" + table.Rows.Count + "  当前：" + i + "行 IDS案件关系  =================");
                sumCount++;
                xfrmProcess.progressBarControl.Position = sumCount;
                xfrmProcess.lbSuccess.Text = sumCount.ToString();
                xfrmProcess.Refresh();
                Application.DoEvents();
                int nCaseID = int.Parse(table.Rows[i]["n_CaseID"].ToString());
                strSql = string.Format(
@"SELECT dbo.TCase_Base.n_CaseID
FROM dbo.TCase_Base  
LEFT JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_CaseID = dbo.TCase_Base.n_CaseID
LEFT JOIN dbo.TPCase_LawInfo ON dbo.TPCase_LawInfo.n_ID = dbo.TPCase_Patent.n_LawID
WHERE 
(ISNULL(dbo.TPCase_LawInfo.s_AppNo,'')<>'' AND dbo.TPCase_LawInfo.s_AppNo IN (SELECT s_PNum FROM dbo.TPCase_Priority where n_CaseID = '{0}'))
OR (ISNULL(dbo.TPCase_LawInfo.s_PCTAppNo,'')<>'' AND (dbo.TPCase_LawInfo.s_PCTAppNo IN (SELECT s_PNum FROM dbo.TPCase_Priority where n_CaseID = '{0}')
OR dbo.TPCase_LawInfo.s_PCTAppNo IN (SELECT TOP 1 s_PCTAppNo FROM dbo.TPCase_LawInfo RIGHT JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_LawID = dbo.TPCase_LawInfo.n_ID WHERE n_CaseID = '{0}')))
UNION
SELECT dbo.TCase_Base.n_CaseID
FROM dbo.TCase_Base  
LEFT JOIN dbo.TPCase_Priority ON dbo.TPCase_Priority.n_CaseID = dbo.TCase_Base.n_CaseID
WHERE ISNULL(s_PNum,'')<>'' AND (s_PNum IN (SELECT s_PNum FROM dbo.TPCase_Priority where n_CaseID = '{0}') 
OR s_PNum = (SELECT TOP 1 s_PCTAppNo FROM dbo.TPCase_LawInfo RIGHT JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_LawID = dbo.TPCase_LawInfo.n_ID WHERE n_CaseID = '{0}')
OR s_PNum = (SELECT TOP 1 s_AppNo FROM dbo.TPCase_LawInfo RIGHT JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_LawID = dbo.TPCase_LawInfo.n_ID WHERE n_CaseID = '{0}'))", nCaseID);
                DataTable relationtable = _dbHelper.GetDataTablebySql(strSql, conn);

                //3、将所有案件与这个美国案建立起IDS关系，注意要在集合中去除该案件本身。 
                var dv = new DataView(relationtable);
                DataTable dt2 = dv.ToTable(true, "n_CaseID");

                for (int k = 0; k < dt2.Rows.Count; k++)
                {
                    int newCaseID = int.Parse(dt2.Rows[k]["n_CaseID"].ToString());
                    if (!nCaseID.Equals(newCaseID))
                    {
                        InsertCodeCaseRelative(newCaseID, nCaseID);
                    }
                }
            }
            if (table.Rows.Count <= 0)
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("未查询到美国案件");
            }
            #endregion

            DateTime end = DateTime.Now;
            TimeSpan ts = end.Subtract(begin).Duration();
            xfrmProcess.progressBarControl.Position = xfrmProcess.progressBarControl.Properties.Maximum;
            xfrmProcess.lbTotalTime.Text = ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒";
            xfrmProcess.lbTotalSum.Text = lostNum + "条";
            Application.DoEvents();
            Thread.Sleep(8000);
            xfrmProcess.Focus();
            xfrmProcess.Close();
            conn.Close();
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============IDS案件关系 End " + ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒" + "=================");
        }

        //创建IDS案件关系
        private void InsertCodeCaseRelative(int caseID, int HKNum)
        {
            const string strSql = "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='IDS' AND s_MasterName='递交方' AND s_SlaveName='接收方' AND s_IPType='P'";
            int n_ID = _dbHelper.GetbySql(strSql, com_databasename.Text, conn);
            if (caseID > 0)
            {
                int NUMS =
                    _dbHelper.GetbySql("SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where ((n_CaseIDA=" + HKNum +
                               " and n_CaseIDB=" + caseID + ") or  (n_CaseIDA=" + caseID + " and n_CaseIDB=" + HKNum + "))and n_CodeRelativeID=" + n_ID, com_databasename.Text, conn);
                if (NUMS <= 0 && HKNum > 0)
                {
                    var _tMainFiles = new T_MainFiles();
                    _tMainFiles.InsertTCaseCaseRelative(HKNum, caseID, n_ID, 0, 0, com_databasename.Text, conn);
                }
            }
            else
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("无法添加");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            IDSCaseRelative();
        }
        #endregion

        #region 4.同族专利

        private void SameRaceCaseRelative()
        {
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============同族专利案件关系 Start =================");
            sumCount = 0;
            conn.Open();
            conn.ChangeDatabase(com_databasename.Text.Trim()); //重新指定数据库 
            DateTime begin = DateTime.Now;

            #region
            //1、从系统中筛选出所有优先权案件。
            string strSql = @"
SELECT dbo.TPCase_Patent.n_CaseID AS 'MainCaseID',dbo.TPCase_Priority.n_CaseID AS 'RelateCaseID'
FROM dbo.TPCase_Priority
INNER JOIN dbo.TPCase_LawInfo ON dbo.TPCase_LawInfo.s_AppNo = dbo.TPCase_Priority.s_PNum 
INNER JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_LawID = dbo.TPCase_LawInfo.n_ID
INNER JOIN dbo.TCase_Base ON dbo.TCase_Base.n_CaseID = dbo.TPCase_Patent.n_CaseID
WHERE n_BusinessTypeID IN (SELECT n_ID FROM dbo.TCode_BusinessType WHERE s_Name <> 'PCT国际阶段申请')";
            DataTable table = _dbHelper.GetDataTablebySql(strSql, conn);

            strSql = @"
WITH PCTCase AS (
SELECT dbo.TCase_Base.n_CaseID,s_PCTAppNo
FROM dbo.TPCase_LawInfo
INNER JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_LawID = dbo.TPCase_LawInfo.n_ID
INNER JOIN dbo.TCase_Base ON dbo.TCase_Base.n_CaseID = dbo.TPCase_Patent.n_CaseID AND dbo.TCase_Base.n_BusinessTypeID IN (SELECT n_ID FROM dbo.TCode_BusinessType WHERE s_Name = 'PCT国际阶段申请') 
WHERE ISNULL(s_PCTAppNo,'') <> ''
),
PCTNCase AS (
SELECT dbo.TCase_Base.n_CaseID,s_PCTAppNo
FROM dbo.TPCase_LawInfo
INNER JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_LawID = dbo.TPCase_LawInfo.n_ID
INNER JOIN dbo.TCase_Base ON dbo.TCase_Base.n_CaseID = dbo.TPCase_Patent.n_CaseID AND dbo.TCase_Base.n_BusinessTypeID IN (SELECT n_ID FROM dbo.TCode_BusinessType WHERE s_Name <> 'PCT国际阶段申请') 
WHERE ISNULL(s_PCTAppNo,'') <> ''
)
SELECT PCTCase.n_CaseID AS 'MainCaseID',PCTNCase.n_CaseID AS 'RelateCaseID'
FROM PCTCase INNER JOIN PCTNCase ON PCTNCase.s_PCTAppNo = PCTCase.s_PCTAppNo";
            DataTable tablePCT = _dbHelper.GetDataTablebySql(strSql, conn);

            table.Merge(tablePCT);



            var listGroupCase = table.Rows.Cast<DataRow>().Select(dr => new { MainCaseID = dr["MainCaseID"].ToString(), RelateCaseID = dr["RelateCaseID"].ToString() }).ToList().GroupBy(k => k.MainCaseID).ToList();
            var xfrmProcess = new ProBar();
            xfrmProcess.progressBarControl.Properties.Maximum = listGroupCase.Count;
            xfrmProcess.progressBarControl.Properties.Minimum = 0;
            xfrmProcess.lbTotalSelected.Text = listGroupCase.Count.ToString();
            xfrmProcess.Show();
            for (int i = 0; i < listGroupCase.Count; i++)
            {
                sumCount++;
                xfrmProcess.progressBarControl.Position = sumCount;
                xfrmProcess.lbSuccess.Text = sumCount.ToString();
                xfrmProcess.Refresh();
                Application.DoEvents();
                var listCase = listGroupCase[i].ToList().Select(c => c.RelateCaseID).ToList();
                listCase.Add(listGroupCase[i].Key);
                listCase = listCase.Distinct().ToList();
                var listCaseCopy = new List<string>();
                listCaseCopy.AddRange(listCase);

                foreach (var exceptCaseID in listCase)
                {
                    foreach (var caseID in listCaseCopy)
                    {
                        if (exceptCaseID == caseID) continue;
                        InsertCodeSameRaceCaseRelative(Convert.ToInt32(exceptCaseID), Convert.ToInt32(caseID));
                    }
                    listCaseCopy.Remove(exceptCaseID);
                }
            }
            if (table.Rows.Count <= 0)
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("未查询到同族案件");
            }
            #endregion

            DateTime end = DateTime.Now;
            TimeSpan ts = end.Subtract(begin).Duration();
            xfrmProcess.progressBarControl.Position = xfrmProcess.progressBarControl.Properties.Maximum;
            xfrmProcess.lbTotalTime.Text = ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒";
            xfrmProcess.lbTotalSum.Text = lostNum + "条";
            Application.DoEvents();
            Thread.Sleep(8000);
            xfrmProcess.Focus();
            xfrmProcess.Close();
            conn.Close();
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============同族案件关系 End " + ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒" + "=================");
        }

        //创建同族案件关系
        private void InsertCodeSameRaceCaseRelative(int caseID, int HKNum)
        {
            const string strSql = "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='同族专利' AND s_MasterName='同族' AND s_SlaveName='同族' AND s_IPType='P'";
            int n_ID = _dbHelper.GetbySql(strSql, com_databasename.Text, conn);
            if (caseID > 0)
            {
                int NUMS =
                    _dbHelper.GetbySql("SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where ((n_CaseIDA=" + HKNum +
                               " and n_CaseIDB=" + caseID + ") or  (n_CaseIDA=" + caseID + " and n_CaseIDB=" + HKNum + "))and n_CodeRelativeID=" + n_ID, com_databasename.Text, conn);
                if (NUMS <= 0 && HKNum > 0)
                {
                    var _tMainFiles = new T_MainFiles();
                    _tMainFiles.InsertTCaseCaseRelative(HKNum, caseID, n_ID, 0, 0, com_databasename.Text, conn);
                }
            }
            else
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("无法添加");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SameRaceCaseRelative();
        }
        #endregion

        #region 5.更新优先权顺序
        private void button1_Click(object sender, EventArgs e)
        {
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============更新优先权顺序 Start =================");
            sumCount = 0;
            conn.Open();
            DateTime begin = DateTime.Now;
            conn.ChangeDatabase(com_databasename.Text.Trim()); //重新指定数据库 
            const string strSql = "SELECT n_CaseID FROM dbo.TCase_Base WHERE n_CaseID IN (SELECT DISTINCT n_CaseID from TPCase_Priority  group by n_CaseID having count(1) >= 2)";
            DataTable table = _dbHelper.GetDataTablebySql(strSql, conn);
            var xfrmProcess = new ProBar();
            xfrmProcess.progressBarControl.Properties.Maximum = table.Rows.Count;
            xfrmProcess.progressBarControl.Properties.Minimum = 0;
            xfrmProcess.lbTotalSelected.Text = table.Rows.Count.ToString();
            xfrmProcess.Show();
            if (table.Rows.Count > 0)
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    sumCount++;
                    xfrmProcess.progressBarControl.Position = sumCount;
                    xfrmProcess.lbSuccess.Text = sumCount.ToString();
                    xfrmProcess.Refresh();
                    Application.DoEvents();
                    int nCaseID = int.Parse(table.Rows[i]["n_CaseID"].ToString());
                    UpdateSeq(nCaseID);
                }
            }
            DateTime end = DateTime.Now;
            TimeSpan ts = end.Subtract(begin).Duration();
            xfrmProcess.progressBarControl.Position = xfrmProcess.progressBarControl.Properties.Maximum;
            xfrmProcess.lbTotalTime.Text = ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒";
            xfrmProcess.lbTotalSum.Text = lostNum + "条";
            Application.DoEvents();
            Thread.Sleep(8000);
            xfrmProcess.Focus();
            xfrmProcess.Close();
            conn.Close();
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============" + lostNum + "条" + ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒" + "=================");
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============更新优先权顺序 End =================");
        }
        //重新排列优先权顺序
        private void UpdateSeq(int nCaseID)
        {
            string strSql = " SELECT  * FROM  TPCase_Priority WHERE n_CaseID=" + nCaseID + " ORDER BY dt_PDate ASC  ";
            DataTable table = _dbHelper.GetDataTablebySql(strSql, conn);
            for (int i = 0; i < table.Rows.Count; i++)
            {
                strSql = "update TPCase_Priority set n_Sequence=" + i + " where n_ID=" + table.Rows[i]["n_ID"];
                _dbHelper.InsertbySql(strSql, i, com_databasename.Text, conn);
            }
        }
        #endregion

        #region 6.更新案件相关关系
        private void button4_Click(object sender, EventArgs e)
        {
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============更新案件相关关系 Start =================");
            sumCount = 0;
            conn.Open();
            DateTime begin = DateTime.Now;
            conn.ChangeDatabase(com_databasename.Text.Trim()); //重新指定数据库 


            string strSql = "SELECT DISTINCT n_CaseIDB FROM TCase_CaseRelative WHERE  s_MasterSlaveRelation=0 AND n_CodeRelativeID in (7,8)";
            DataTable table = _dbHelper.GetDataTablebySql(strSql, conn);
            var xfrmProcess = new ProBar();
            xfrmProcess.progressBarControl.Properties.Maximum = table.Rows.Count;
            xfrmProcess.progressBarControl.Properties.Minimum = 0;
            xfrmProcess.lbTotalSelected.Text = table.Rows.Count.ToString();
            xfrmProcess.Show();
            if (table.Rows.Count > 0)
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    sumCount++;
                    xfrmProcess.progressBarControl.Position = sumCount;
                    xfrmProcess.lbSuccess.Text = sumCount.ToString();
                    xfrmProcess.Refresh();
                    Application.DoEvents();
                    int nCaseIDA = int.Parse(table.Rows[i]["n_CaseIDB"].ToString());

                    strSql = "SELECT n_CaseIDA FROM TCase_CaseRelative WHERE  s_MasterSlaveRelation=0 AND n_CodeRelativeID in (7,8) and n_CaseIDB=" + nCaseIDA;
                    DataTable newtable = _dbHelper.GetDataTablebySql(strSql, conn);
                    for (int j = 0; j < newtable.Rows.Count; j++)
                    {
                        int nCaseIDB = int.Parse(newtable.Rows[j]["n_CaseIDA"].ToString());
                        for (int k = 0; k < newtable.Rows.Count; k++)
                        {
                            nCaseIDA = int.Parse(newtable.Rows[k]["n_CaseIDA"].ToString());
                            if (!nCaseIDA.Equals(nCaseIDB))
                            {
                                InsertTCaseCaseRelative(nCaseIDA, nCaseIDB);
                            }
                        }
                    }
                }
            }


            DateTime end = DateTime.Now;
            TimeSpan ts = end.Subtract(begin).Duration();
            xfrmProcess.progressBarControl.Position = xfrmProcess.progressBarControl.Properties.Maximum;
            xfrmProcess.lbTotalTime.Text = ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒";
            xfrmProcess.lbTotalSum.Text = lostNum + "条";
            Application.DoEvents();
            Thread.Sleep(8000);
            xfrmProcess.Focus();
            xfrmProcess.Close();
            conn.Close();
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============" + lostNum + "条" + ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒" + "=================");
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============更新案件相关关系 End =================");
        }
        private void InsertTCaseCaseRelative(int caseIDA, int caseIDB)
        {
            string strSql =
                     "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='相关案件' AND s_MasterName='' AND s_SlaveName='' AND s_IPType='P'";

            int n_ID = _dbHelper.GetbySql(strSql, com_databasename.Text, conn);
            if (n_ID > 0)
            {
                strSql =
                   "SELECT count(*) as Sunm FROM dbo.TCase_CaseRelative WHERE ((n_CaseIDA=" + caseIDA + " AND n_CaseIDB=" + caseIDB + ") OR (n_CaseIDA=" + caseIDB + " AND n_CaseIDB=" + caseIDA + ") ) AND s_MasterSlaveRelation=0 AND n_CodeRelativeID=" + n_ID;

                int sum = _dbHelper.GetbySql(strSql, com_databasename.Text, conn);

                if (sum <= 0)
                {
                    strSql =
                        " INSERT INTO  dbo.TCase_CaseRelative ( n_CaseIDA ,  n_CaseIDB , dt_CreateDate , dt_EditDate , s_MasterSlaveRelation , n_CodeRelativeID )" +
                        " VALUES  ( " + caseIDB + " , " + caseIDA + " ,  GETDATE() , GETDATE() , 0,  " + n_ID + ")";
                    _dbHelper.InsertbySql(strSql, 0, com_databasename.Text, conn);
                }
            }
        }
        #endregion

        #region 更新官方收据抬头
        private void button7_Click(object sender, EventArgs e)
        {
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============更新申请人官方收据抬头 Start =================");
            sumCount = 0;
            conn.Open();
            DateTime begin = DateTime.Now;
            conn.ChangeDatabase(com_databasename.Text.Trim()); //重新指定数据库 


            const string strSql = "SELECT DISTINCT n_CaseID FROM  dbo.TCase_Base";
            DataTable table = _dbHelper.GetDataTablebySql(strSql, conn);
            var xfrmProcess = new ProBar();
            xfrmProcess.progressBarControl.Properties.Maximum = table.Rows.Count;
            xfrmProcess.progressBarControl.Properties.Minimum = 0;
            xfrmProcess.lbTotalSelected.Text = table.Rows.Count.ToString();
            xfrmProcess.Show();
            var pList = new Dictionary<String, String>();
            if (table.Rows.Count > 0)
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    string nCaseID = table.Rows[i]["n_CaseID"].ToString();
                    string strSql1 = "SELECT s_PayFeePerson,b.s_Name FROM dbo.TCase_Applicant a  LEFT JOIN dbo.TCstmr_Applicant  b ON a.n_ApplicantID=b.n_AppID WHERE  n_CaseID=" + nCaseID + " order by n_Sequence asc ";
                    DataTable newtable = _dbHelper.GetDataTablebySql(strSql1, conn);

                    string value = string.Empty;
                    for (int j = 0; j < newtable.Rows.Count; j++)
                    {
                        value += newtable.Rows[j]["s_Name"] + ":" + newtable.Rows[j]["s_PayFeePerson"]+";";
                    }
                    value = value.TrimEnd(';');
                    if (pList.ContainsKey(nCaseID) == false && !string.IsNullOrEmpty(value))
                    {
                        pList.Add(nCaseID, value);
                    } 
                } 
            }
            string Sql = "";
            foreach (var dic in pList)
            {
                string[] arry = dic.Value.Split(';');
                for (int i = 0; i < arry.Length; i++)
                {
                    if (!string.IsNullOrEmpty(arry[i]))
                    {
                        string[] ar = arry[i].Split(':');
                        if (ar[1].ToString().Equals("Y"))
                        {
                            Sql += "update TCase_Base set s_ElecPayer='" + ar[0] + "' where n_CaseID=" + dic.Key + "              \r\n ";
                            break; ;
                        }
                    }
                }
            }
            _dbHelper.InsertbySql(Sql, 0, "更新申请人官方收据抬头", conn);

            var end = DateTime.Now;
            TimeSpan ts = end.Subtract(begin).Duration();
            xfrmProcess.progressBarControl.Position = xfrmProcess.progressBarControl.Properties.Maximum;
            xfrmProcess.lbTotalTime.Text = ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒";
            xfrmProcess.lbTotalSum.Text = lostNum + "条";
            Application.DoEvents();
            Thread.Sleep(8000);
            xfrmProcess.Focus();
            xfrmProcess.Close();
            conn.Close();
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============" + lostNum + "条" + ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒" + "=================");
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============更新申请人官方收据抬头 End =================");
        }

        public static void DicSample1(string nCaseID, string value,Dictionary<String, String> pList)
        {

         
            try
            {
              
            }
            catch (System.Exception e)
            {
                Console.WriteLine("Error: {0}", e.Message);
            }
        }
        #endregion

        #endregion
    }
}
