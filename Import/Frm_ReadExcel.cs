using System.Globalization;
using Import;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using DevExpress.Data.PLinq.Helpers;

namespace Winform_SqlBulkCopy
{
    public partial class Frm_ReadExcel : Form
    {

        #region 全局变量

        private DataSet ds;
        private string[] tablenames;
        private SqlConnection conn;
        private readonly string connstr;
        private int OutFileID;
        private int InFileID;
        private List<SqlBulkCopyColumnMapping> sqlBulkCopyparameters;

        #endregion

        #region 构造函数

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="_connstr">sqlserver连接字符串</param>
        public Frm_ReadExcel(string _connstr)
        {
            connstr = _connstr;
            InitializeComponent();
            Load += new EventHandler(Frm_ReadExcel_Load);
        }

        #endregion

        #region 窗体加载

        private void Frm_ReadExcel_Load(object sender, EventArgs e)
        {
            Init();
            EventHand();
        }

        #endregion

        #region 初始化

        private void Init()
        {
            conn = new SqlConnection(connstr); //SqlConnection实例化
            MaximizeBox = false; //禁用最小化
            MaximumSize = MinimumSize = Size; //固定当前大小
            txt_filepath.ReadOnly = true;
            com_databasename.DropDownStyle = com_tablename.DropDownStyle = ComboBoxStyle.DropDownList; //下拉框只可选

            this.cbxSheet.DropDownStyle = cbxSheet.DropDownStyle = ComboBoxStyle.DropDownList; //下拉框只可选

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

        #endregion

        #region 控件事件挂接

        private void EventHand()
        {
            bt_see.Click += new EventHandler(bt_see_Click);
            bt_ok.Click += new EventHandler(bt_ok_Click);
            FormClosing += new FormClosingEventHandler(Frm_ReadExcel_FormClosing);
        }

        #endregion

        #region 修改文件类型

        /// <summary>
        /// 修改文件类型
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void com_FileType_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //if (xradioRegOnline.EditValue.Equals("任务时限"))
                //{
                //    label7.Visible = true;
                //    button1.Visible = true;
                //    //bt_SetColumns.Visible = true;
                //    //bt_instruction.Visible = true;
                //}
                //else
                //{
                //    label7.Visible = false;
                //    button1.Visible = false;
                //    //bt_SetColumns.Visible = false;
                //    //bt_instruction.Visible = false;
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #endregion

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

        #region 委托、获取参数

        /// <summary>
        /// 委托、获取参数
        /// </summary>
        //private void GetColumns(List<SqlBulkCopyColumnMapping> _SqlBulkCopyparameters)
        //{
        //    SqlBulkCopyColumnMapping m1 = new SqlBulkCopyColumnMapping("ShopCode", "ShopName");
        //    _SqlBulkCopyparameters.Add(m1);
        //    SqlBulkCopyparameters = _SqlBulkCopyparameters;
        //}

        //private void NewColumns()
        //{

        //    SqlBulkCopyparameters = new List<SqlBulkCopyColumnMapping>();
        //    SqlBulkCopyColumnMapping m1 = new SqlBulkCopyColumnMapping("我方卷号", "HKNum");
        //    SqlBulkCopyparameters.Add(m1);

        //    SqlBulkCopyColumnMapping m2 = new SqlBulkCopyColumnMapping("优先权号", "Propiry");
        //    SqlBulkCopyparameters.Add(m2);

        //    SqlBulkCopyColumnMapping m3 = new SqlBulkCopyColumnMapping("优先权号", "ProDate");
        //    SqlBulkCopyparameters.Add(m3);

        //    SqlBulkCopyColumnMapping m4 = new SqlBulkCopyColumnMapping("优先权国家", "ProCountry");
        //    SqlBulkCopyparameters.Add(m4);
        //}

        #endregion

        private int sumCount = 0;
        private int num = 0;

        #region 导入数据

        private void getType()
        {
            conn.Open();
            conn.ChangeDatabase(com_databasename.Text.Trim()); //重新指定数据库 
            string strSql = "select OID FROM dbo.XPObjectType WHERE TypeName LIKE 'DataEntities.Element.Files.OutFile'";
            string strSql2 = "select OID FROM dbo.XPObjectType WHERE TypeName LIKE 'DataEntities.Element.Files.InFile'";
            OutFileID = GetIDbySql(strSql);
            InFileID = GetIDbySql(strSql2);
            conn.Close();
        }

        private void bt_ok_Click(object sender, EventArgs e)
        {
            getType();
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("-------------------------------------" +
                                                                       xradioRegOnline.EditValue +
                                                                       "  Start -------------------------------------");
            string mgr = string.Empty;

            var xfrmProcess = new ProBar();
            try
            {
                DateTime begin = DateTime.Now;
                //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("开始时间:" + begin);
                conn.Open();
                conn.ChangeDatabase(com_databasename.Text.Trim()); //重新指定数据库 
                DataTable table = (DataTable)dgv_show.DataSource; //获取要Copy的数据源
                //s_IOType I：来文 O：发文 T：其它文件 
                //s_ClientGov  C: 客户  O: 官方 
                sumCount = 0;
                xfrmProcess.progressBarControl.Properties.Maximum = table.Rows.Count;
                xfrmProcess.progressBarControl.Properties.Minimum = 0;
                xfrmProcess.lbTotalSelected.Text = table.Rows.Count.ToString();
                xfrmProcess.Show();

                #region 1.国内-收文数据导入

                if (xradioRegOnline.EditValue.ToString() == "国内-收文数据导入") //T_MainFiles
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int ResultNum = InsertFileIn(i, table.Rows[i]);
                        if (ResultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 2.国内-专利数据补充导入

                else if (xradioRegOnline.EditValue.ToString() == "国内-专利数据补充导入")
                {
                    #region 国内-专利数据补充导入

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int ResultNum = InsertPatented(i, table.Rows[i]);
                        if (ResultNum == 0)
                        {
                            num++;
                        }
                    }

                    #endregion
                }
                #endregion

                #region 3.国外-法律信息及日志表

                else if (xradioRegOnline.EditValue.ToString() == "国外-法律信息及日志表")
                {
                    #region 国外-法律信息及日志表

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int ResultNum = TCaseLaw(i, table.Rows[i]);
                        if (ResultNum == 0)
                        {
                            num++;
                        }
                    }

                    #endregion
                }
                #endregion

                #region 4.国内-OA数据补充导入

                else if (xradioRegOnline.EditValue.ToString() == "国内-OA数据补充导入")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int resultNum = OA(i, table.Rows[i]);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 5.实体信息表

                else if (xradioRegOnline.EditValue.ToString() == "实体信息表")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int ResultNum = TCaseApplicant(i, table.Rows[i]);
                        if (ResultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 6.国外-国外库发明人表

                else if (xradioRegOnline.EditValue.ToString() == "国外-国外库发明人表")
                {
                   
                    #region 国外库发明人表

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int ResultNum = TPCaseInventor(i, table.Rows[i]);
                        if (ResultNum == 0)
                        {
                            num++;
                        }
                    }

                    #endregion
                }
                #endregion

                #region 7.国外-国外库时限备注表

                else if (xradioRegOnline.EditValue.ToString() == "国外-国外库时限备注表")
                {
                    #region 国外-国外库时限备注表

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int RsultNum = Case_Memo(i, table.Rows[i]);
                        if (RsultNum == 0)
                        {
                            num++;
                        }
                    }

                    #endregion
                }
                #endregion

                #region 8.任务时限

                else if (xradioRegOnline.EditValue.ToString() == "任务时限")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int ResultNum = ImportTask(i, table.Rows[i]);
                        if (ResultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 9.香港-专利数据补充导入

                else if (xradioRegOnline.EditValue.ToString() == "香港-专利数据补充导入")
                {
                    //DataTable dt = new DataTable();
                    //克隆表结构
                    //dt = table.Clone();
                    //string strSql = string.Empty;

                    #region 香港-专利数据补充导入

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int Result = HongKang(i, table.Rows[i]);
                        if (Result == 0)
                        {
                            num++;
                        }
                    }

                    #endregion
                }
                #endregion

                #region 10.年费

                else if (xradioRegOnline.EditValue.ToString() == "年费")
                {
                    #region 年费
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int resultNum = InsertFee(i, table.Rows[i]);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }

                    #endregion
                }
                #endregion

                #region  11.案件处理人

                else if (xradioRegOnline.EditValue.ToString() == "案件处理人")
                {
                    #region 案件处理人

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString(CultureInfo.InvariantCulture);
                        xfrmProcess.Refresh();
                        int result = InsertCaseAttorney(table, i, table.Rows[i]);
                        if (result == 0)
                        {
                            num++;
                        }
                    }

                    #endregion
                }
                #endregion

                #region  12.客户要求

                else if (xradioRegOnline.EditValue.ToString() == "客户要求")
                {
                    #region 客户要求

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString(CultureInfo.InvariantCulture);
                        xfrmProcess.Refresh();
                        int result = InsertDemandClient(table.Rows[i],i);
                        if (result == 0)
                        {
                            num++;
                        }
                    }

                    #endregion
                }
                #endregion

                #region 13.优先权

                else if (xradioRegOnline.EditValue.ToString() == "优先权")
                {
                    #region 优先权

                    DataRow[] NewTable = table.Select("", "优先权号  ");
                    for (int i = 0; i < NewTable.Length; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int ReturnNum = TPCasePriority(i, NewTable[i]);
                        if (ReturnNum == 0)
                        {
                            num++;
                        }
                    }

                    #endregion
                }
                #endregion

                #region 14.澳门案件-澳门延伸

                else if (xradioRegOnline.EditValue.ToString() == "澳门案件-澳门延伸")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString(CultureInfo.InvariantCulture);
                        xfrmProcess.Refresh();
                        int result = InsertMacaoApplication(table.Rows[i],i);
                        if (result == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 15.相关案件-双申

                else if (xradioRegOnline.EditValue.ToString() == "相关案件-双申")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString(CultureInfo.InvariantCulture);
                        xfrmProcess.Refresh();
                        int result = InsertDoubleShen(table.Rows[i],i);
                        if (result == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 16.递交机构配置

                else if (xradioRegOnline.EditValue.ToString() == "递交机构")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString(CultureInfo.InvariantCulture);
                        xfrmProcess.Refresh();

                        int resultNum = InserIntoOrgin(table.Rows[i], i);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 17.根据业务类型修改申请方式

                else if (xradioRegOnline.EditValue.ToString() == "申请方式")
                {
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
                                int resultNum = UpdateUSAType(table.Rows[i], i);
                                if (resultNum == 0)
                                {
                                    num++;
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

                            int resultNum = UpdateType(table.Rows[i], i);
                            if (resultNum == 0)
                            {
                                num++;
                            }
                        }
                    }
                }
                #endregion

                #region 18.去案所
                else if (xradioRegOnline.EditValue.ToString() == "案件选择")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = UpdateManualCreateChain(table.Rows[i], i);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 19.案件部门导入
                else if (xradioRegOnline.EditValue.ToString() == "部门核对")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = UpdateOrg(table.Rows[i], i);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 20.申请人需要导入电子缴费单缴费人
                else if (xradioRegOnline.EditValue.ToString() == "电子缴费单缴费人")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = UpdateTCstmrApplicant(table.Rows[i], i);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 21.相关客户
                else if (xradioRegOnline.EditValue.ToString() == "相关客户")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = UpdateRelatedcustomers(table.Rows[i], i);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 22.客户要求代码
                else if (xradioRegOnline.EditValue.ToString() == "客户要求代码")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = InsertDemand(table.Rows[i], i);//增加客户、申请人、客户-申请人要求
                        InsertCaseDemand(table.Rows[i]); //拷贝案子要求
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                    sumCount = 0;
                    DataRow[] newtableRow = table.Select(" [客户ID] is not null and [申请人ID]  is not null");
                    for (int i = 0; i < newtableRow.Length; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();
                        int resultNum = DeleteDemand(newtableRow[i]);//
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 25.分案信息
                else if (xradioRegOnline.EditValue.ToString() == "分案信息")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = UpdateCaseDivisionInfo(table.Rows[i]);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 
                #region 23.总委托书号
                else if (xradioRegOnline.EditValue.ToString() == "总委托书号")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = UpdateTotalCommissionNumber(table.Rows[i], i);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 24.初审合格日 已注释
                //else if (xradioRegOnline.EditValue.ToString() == "初审合格日")
                //{
                //    for (int i = 0; i < table.Rows.Count; i++)
                //    {
                //        sumCount++;
                //        xfrmProcess.progressBarControl.Position = sumCount;
                //        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                //        xfrmProcess.Refresh();

                //        int resultNum = UpdateTPCaseLawInfo(table.Rows[i], i);
                //        if (resultNum == 0)
                //        {
                //            num++;
                //        }
                //    }
                //}
                #endregion

                #region 26.案件状态 注释
                //else if (xradioRegOnline.EditValue.ToString() == "案件状态")
                //{
                //    for (int i = 0; i < table.Rows.Count; i++)
                //    {
                //        sumCount++;
                //        xfrmProcess.progressBarControl.Position = sumCount;
                //        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                //        xfrmProcess.Refresh();

                //        int resultNum = UpdateCasesCaseStatus(table.Rows[i], i);
                //        if (resultNum == 0)
                //        {
                //            num++;
                //        }
                //    }
                //}
                #endregion

                #endregion

                #region 27.香港优先权
                else if (xradioRegOnline.EditValue.ToString() == "优先权（香港）")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = HongKangPriority(i, table.Rows[i]);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 28.控制要求
                else if (xradioRegOnline.EditValue.ToString() == "控制要求")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = InsertCodeDemandType(table.Rows[i]);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 29.客户信息
                else if (xradioRegOnline.EditValue.ToString() == "客户信息")
                {
                    string _selectValue = cbxSheet.SelectedValue.ToString();
                    if (_selectValue.Contains("联系人"))
                    {
                        DataRow[] newtableRow = table.Select(" [客户代码] is not null");
                        for (int i = 0; i < newtableRow.Length; i++)
                        {
                            sumCount++;
                            xfrmProcess.progressBarControl.Position = sumCount;
                            xfrmProcess.lbSuccess.Text = sumCount.ToString();
                            xfrmProcess.Refresh();
                            int resultNum = AddClientContact(newtableRow[i], i);//
                            if (resultNum == 0)
                            {
                                num++;
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
                                resultNum = InsertTCstmrClient(table.Rows[i], i);
                            }
                            else if (_selectValue.Contains("地址"))
                            {
                                resultNum = AddClientAddress(table.Rows[i], i);
                            }
                            //else if (_selectValue.Contains("联系人"))
                            //{
                            //    resultNum = AddClientContact(table.Rows[i], i);
                            //}
                            else if (_selectValue.Contains("财务信息"))
                            {
                                resultNum = AddBill(table.Rows[i]);
                            }
                            if (resultNum == 0)
                            {
                                num++;
                            }
                        }
                    }
                    updateApplicantandClient();
                }
                #endregion

                #region 30.申请人信息
                else if (xradioRegOnline.EditValue.ToString() == "申请人信息")
                {
                    string _selectValue = cbxSheet.SelectedValue.ToString();
                    if (_selectValue.Contains("联系人"))
                    {
                        DataRow[] newtableRow = table.Select(" [申请人代码] is not null");
                        for (int i = 0; i < newtableRow.Length; i++)
                        {
                            sumCount++;
                            xfrmProcess.progressBarControl.Position = sumCount;
                            xfrmProcess.lbSuccess.Text = sumCount.ToString();
                            xfrmProcess.Refresh();
                            int resultNum = AddAppContact(newtableRow[i], i);//
                            if (resultNum == 0)
                            {
                                num++;
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
                                resultNum = InsertTCstmrApplicant(table.Rows[i], i);
                            }
                            else if (_selectValue.Contains("地址"))
                            {
                                resultNum = AddAppAddress(table.Rows[i], i);
                            }
                            //else if (_selectValue.Contains("联系人"))
                            //{
                            //    resultNum = AddAppContact(table.Rows[i], i);
                            //}
                            else if (_selectValue.Contains("财务信息"))
                            {
                                resultNum = AddAppBill(table.Rows[i]);
                            }
                            if (resultNum == 0)
                            {
                                num++;
                            }
                        }
                    }
                    updateApplicantandClient();
                }
                #endregion

                #region 31.外代理信息
                else if (xradioRegOnline.EditValue.ToString() == "外代理信息")
                {
                    string _selectValue = cbxSheet.SelectedValue.ToString();
                    if (_selectValue.Contains("联系人"))
                    {
                        DataRow[] newtableRow = table.Select(" [外代理] is not null");
                        for (int i = 0; i < newtableRow.Length; i++)
                        {
                            sumCount++;
                            xfrmProcess.progressBarControl.Position = sumCount;
                            xfrmProcess.lbSuccess.Text = sumCount.ToString();
                            xfrmProcess.Refresh();
                            int resultNum = AddAgencyContact(newtableRow[i], i);//
                            if (resultNum == 0)
                            {
                                num++;
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
                                resultNum = InsertTCstmrCoopAgency(table.Rows[i], i);
                            }
                            else if (_selectValue.Contains("地址"))
                            {
                                resultNum = AddAgencyAddress(table.Rows[i], i);
                            }
                            //else if (_selectValue.Contains("联系人"))
                            //{
                            //    resultNum = AddAgencyContact(table.Rows[i], i);
                            //}
                            else if (_selectValue.Contains("财务信息"))
                            {
                                resultNum = AddAgencyBill(table.Rows[i]);
                            }
                            if (resultNum == 0)
                            {
                                num++;
                            }
                        }
                    }
                }
                #endregion

                #region 自定义属性
                else if (xradioRegOnline.EditValue.ToString() == "自定义属性")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = InsertCodeCaseCustomField(table.Rows[i],i);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 要求/发文
                else if (xradioRegOnline.EditValue.ToString() == "要求/发文")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = sGetType(table.Rows[i], i);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 备注
                else if (xradioRegOnline.EditValue.ToString() == "备注")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = InsertTCaseMemo(table.Rows[i], i);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion

                #region 档案位置
                else if (xradioRegOnline.EditValue.ToString() == "档案位置")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = InsertFileLocation(table.Rows[i], i);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion 

                #region 更新申请人译名
                else if (xradioRegOnline.EditValue.ToString() == "更新申请人译名")
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        sumCount++;
                        xfrmProcess.progressBarControl.Position = sumCount;
                        xfrmProcess.lbSuccess.Text = sumCount.ToString();
                        xfrmProcess.Refresh();

                        int resultNum = UpdateApplicant(table.Rows[i], i);
                        if (resultNum == 0)
                        {
                            num++;
                        }
                    }
                }
                #endregion 

                #region 结束

                else //整体不需要转换导入
                {
                    SqlBulkCopy bulkcopy = new SqlBulkCopy(conn);
                    foreach (SqlBulkCopyColumnMapping columnsmapping in sqlBulkCopyparameters)
                        bulkcopy.ColumnMappings.Add(columnsmapping); //加载文件与表的列名 
                    bulkcopy.DestinationTableName = com_tablename.Text.Trim(); //指定目标表的表名
                    bulkcopy.WriteToServer(table); //将table Copy到数据库 
                }
                DateTime end = DateTime.Now;
                TimeSpan ts = end.Subtract(begin).Duration();
                xfrmProcess.progressBarControl.Position = xfrmProcess.progressBarControl.Properties.Maximum;
                xfrmProcess.lbTotalTime.Text = ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒";
                xfrmProcess.lbTotalSum.Text = num + "条";

                Application.DoEvents();
                Thread.Sleep(8000);
                xfrmProcess.Focus();
                xfrmProcess.Close();
                //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("结束时间:" + end + "===" + ts);
                mgr = string.Format("服务器:{0}\n数据库:{1}\n共复制{2}条数据\n{3}条数据未找到“我方卷号”\n耗时{4}天{5}小时{6}分{7}秒",
                                    conn.DataSource, com_databasename.Text, table.Rows.Count, num, ts.Days, ts.Hours,
                                    ts.Minutes, ts.Seconds);
                conn.Close();

                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(mgr);

                #endregion
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

        #endregion

        #region 导入数据方法

        #region 2016-12-07新增

        #region 自定义属性与案件关系
        private int InsertCodeCaseCustomField(DataRow dataRow, int row)
        {
            if (dataRow["自定义属性名称"] != null && !string.IsNullOrEmpty(dataRow["自定义属性名称"].ToString()))
            { 
                string sNo = dataRow["我方卷号"].ToString().Trim();
                int numHk = GetIDbyName(sNo, 2);
                if (numHk > 0)
                {
                    string strSql =
                        " SELECT n_ID FROM TCode_CaseCustomField WHERE  s_IPType='P' AND s_IsActive='Y' AND s_CustomFieldName IN ('" +
                        dataRow["自定义属性名称"] + "')";
                    int nCaseFieldID = GetIDbySql(strSql);
                    if (nCaseFieldID > 0)
                    {
                        strSql = " SELECT n_ID FROM TCase_CaseCustomField WHERE n_CaseID=" + numHk +
                                 " AND n_FieldCodeID=" + nCaseFieldID;
                        int nID = GetIDbySql(strSql);
                        if (nID > 0)
                        {
                            strSql = "update TCase_CaseCustomField set s_Value='" + dataRow["自定义属性值"] +
                                     "' WHERE n_CaseID=" +
                                     numHk + " AND n_FieldCodeID=" + nCaseFieldID;
                        }
                        else
                        {
                            strSql =
                                "INSERT INTO dbo.TCase_CaseCustomField( n_CaseID, n_FieldCodeID, s_Value ) VALUES  (" +
                                numHk + "," + nCaseFieldID + "," + dataRow["自定义属性值"] + ")";
                        }
                    }
                    else
                    {
                        InsertbySql(
                            "INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" +
                            numHk + "," + row + ",'" + sNo + "','未找到“我方卷号”为:" + sNo +
                            "','InsertCodeCaseCustomField','未找到自定义属性-" + row + "')");
                    }
                   return InsertbySql(strSql);
                }else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + numHk + "," + row + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','InsertCodeCaseCustomField','n_ID：未找到我方卷号-" + row + "')"); 
                }
            } 
            return 0;
        }

        #endregion

        #region 发文官方/发文客户/案件个案要求
        private int sGetType(DataRow dataRow, int row)
        {
            if (dataRow["栏目名称"] != null && dataRow["我方文号"]!=null)
            {
                if (dataRow["导入形式"] != null &&
                    (dataRow["导入形式"].ToString().Equals("发客户文") || dataRow["导入形式"].ToString().Equals("发官方文")))
                {
                    return InsertOutFile(dataRow, row);
                }
                else if (dataRow["导入形式"] != null && dataRow["导入形式"].ToString().Equals("案件客户要求"))
                {
                    return InsertOnlyCaseDemand(dataRow, row);
                }
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(0," + row + ",'" + dataRow["我方文号"] + "','未找到“我方卷号”为:" + dataRow["我方文号"] + "','InsertCodeCaseCustomField','文件往来-" + row + "')"); 
            }
            return 0;
        }

        private int InsertOutFile(DataRow dataRow, int row)
        {
            if (dataRow["我方卷号"]!=null && dataRow["导入形式"] != null)
            {
                string sNo = dataRow["我方卷号"].ToString().Trim();
                int numHk = GetIDbyName(sNo, 2);
                if (numHk <= 0)
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + numHk + "," + row + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','InsertCodeCaseCustomField','n_ID：未找到我方卷号：文件往来-" + row + "')");
                    return 0;
                }
                else
                {
                    string sClientGov = "O"; //C: 客户 O: 官方  s_IOType I：来文 O：发文 T：其它文件 

                    if (dataRow["导入形式"] != null && dataRow["导入形式"].ToString().Equals("发客户文"))
                    {
                        sClientGov = "C";
                    }
                    string content = dataRow["栏目名称"].ToString() + "：" + dataRow["栏目内容"].ToString();
                    DateTime insertTime = DateTime.Now;
                    string strSql =
                        "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,ObjectType,dt_EditDate,s_Name,dt_CreateDate,s_ClientGov,s_IOType,s_Status,s_Abstact)" +
                        "VALUES  ('发文官方/发文客户/案件个案要求-" + row + "'," + OutFileID + ",'" + insertTime + "','" +
                        content.Replace("'", "''") + "','" + insertTime + "','" + sClientGov + "','O','Y','" + dataRow["备注"].ToString() + "')";

                    string sqlS =
                        "SELECT top 1 n_FileID FROM dbo.T_MainFiles WHERE  s_IOType='O' AND s_Status='Y' and  s_ClientGov='" +
                        sClientGov + "' and s_Name='" + content.Replace("'", "''") + "' and ObjectType=" + OutFileID +
                        " and dt_CreateDate='" + insertTime + "' order by n_FileID desc ";

                    int nFileID = 0;
                    if (InsertbySql(strSql, row) > 0)
                    {
                        nFileID = GetIDbySql(sqlS);
                    }
                    if (nFileID > 0)
                    {
                        string sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_FileInCase WHERE n_FileID=" + nFileID +
                                     " and n_CaseID=" + numHk;
                        int sumFileInCase = GetIDbySql(sql);
                        if (sumFileInCase <= 0)
                        {
                            sql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                                  "VALUES  (" + numHk + " ," + nFileID + ",'Y')";
                            sumFileInCase = InsertbySql(sql, row); //记录文件、案件与程序的关系表 
                            if (sumFileInCase <= 0)
                            {
                                InsertbySql(
                                    "INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" +
                                    numHk + "," + row + ",0,'T_FileInCase插入数据错误','T_FileInCase','发文官方/发文客户/案件个案要求-" +
                                    row +
                                    "','" + sql.Replace("'", "''") + "')");
                            }
                        }
                        else
                        {
                            InsertbySql(
                                "INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" +
                                numHk + "," + row + ",0,'已存在案件和文件关系，无需插入','T_FileInCase','发文官方/发文客户/案件个案要求-" + row +
                                "','" +
                                sql.Replace("'", "''") + "')");
                        }
                        int nGovOfficeID = 0;
                        if (sClientGov.Equals("O")) //s_ClientGov C: 客户 O: 官方
                        {
                            nGovOfficeID = 21; //中国国家知识产权局
                        }
                        if (sumFileInCase > 0)
                        {
                            sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_OutFiles WHERE n_FileID=" + nFileID;
                            int sumcount = GetIDbySql(sql);
                            if (sumcount <= 0)
                            {
                                sqlS =
                                    "INSERT INTO dbo.T_OutFiles( n_FileID ,n_CheckedOutBy , n_GovOfficeID , s_FileStatus, dt_StatusDate ,dt_WriteDate ," +
                                    "n_WriterID , n_SubmiterID ,  n_PrintNum , n_PageNum ,n_ReFileID  ,n_Count ,s_FileType ,n_LatestCheckInfoID)" +
                                    "VALUES  (" + nFileID + ",0 ," + nGovOfficeID + " ,'W' ,'" + DateTime.Now + "' ,'" +
                                    DateTime.Now + "',0 ,0 ,1,0 ,0,0 ,'new',0 )";

                                int outFiles = InsertbySql(sqlS, row);
                                if (outFiles <= 0)
                                {
                                    InsertbySql(
                                        "INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" +
                                        numHk + "," + row + ",0,'T_OutFiles插入数据错误','T_OutFiles','发文官方/发文客户/案件个案要求-" +
                                        row +
                                        "','" + sqlS.Replace("'", "''") + "')");
                                }
                            }
                        }
                    }
                    return 1;
                }
            }
            return 0;
        }

        private int InsertOnlyCaseDemand(DataRow dataRow, int row)
        {
            string sNo = dataRow["我方卷号"].ToString().Trim();
            int numHk = GetIDbyName(sNo, 2);
            string content = dataRow["栏目名称"].ToString().Replace("'", "''") + "：" + dataRow["栏目内容"].ToString().Replace("'", "''");
            string Sql = "SELECT n_ID  FROM T_Demand where  s_Title='" + content + "' and n_CaseID=" + numHk + " and s_Description='" + dataRow["备注"].ToString().Replace("'", "''") + "'";
            int n_ID = GetIDbySql(Sql);
            if (n_ID <= 0 && numHk>0)
            {
                Sql = "INSERT INTO dbo.T_Demand(s_sourcetype1,dt_EditDate,s_Title,dt_CreateDate,n_CaseID,s_Description)" +
                      "VALUES  ('案件要求','" + DateTime.Now + "','" + content + "','" + DateTime.Now + "'," + numHk + ",'" + dataRow["备注"].ToString().Replace("'", "''") + "')";
                int ResultNum = InsertbySql(Sql, row);
                if (ResultNum.Equals(0))
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + numHk + "," + row + ",'','增加案件要求失败[InsertOnlyCaseDemand]','T_Demand','发文记录和案件要求-" + row + "','" + Sql.Replace("'", "''") + "')");
                }
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + numHk + "," + row + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','InsertOnlyCaseDemand','n_ID：未找到我方卷号或者已存在案件要求-" + row + "')");
            }
            return 0;
        }
        #endregion

        #region 备注
        private int InsertTCaseMemo(DataRow dataRow, int row)
        {
            string sNo = dataRow["我方卷号"].ToString().Trim();
            int numHk = GetIDbyName(sNo, 2);
            string Sql = "SELECT n_ID  FROM TCase_Memo where  s_Memo='" + dataRow["备注内容"].ToString().Replace("'", "''") + "' and n_CaseID=" + numHk + " and s_Type='" + dataRow["备注类型"].ToString().Replace("'", "''") + "'";
            int n_ID = GetIDbySql(Sql.Replace("'", "''"));
            if (numHk>0)
            {
                if (n_ID <= 0)
                {
                    if (!string.IsNullOrEmpty(dataRow["备注内容"].ToString()) ||
                        !string.IsNullOrEmpty(dataRow["备注类型"].ToString()))
                    {
                        Sql = "INSERT INTO  dbo.TCase_Memo(n_CaseID ,s_Memo ,s_Type ,dt_CreateDate ,dt_EditDate)" +
                              "VALUES  (" + numHk + ",'" + dataRow["备注内容"].ToString().Replace("'", "''") + "','" +
                              dataRow["备注类型"].ToString().Replace("'", "''") + "','" + DateTime.Now +
                              "','" + DateTime.Now + "')";
                        int ResultNum = InsertbySql(Sql.Replace("'", "''"), row);
                        if (ResultNum.Equals(0))
                        {
                            return
                                InsertbySql(
                                    "INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" +
                                    numHk + "," + row +
                                    ",'','增加备注失败[InsertTCaseMemo]','TCase_Memo','增加备注失败[InsertTCaseMemo]','" +
                                    Sql.Replace("'", "''") + "')");
                        }
                    }
                }
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + numHk + "," + row + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','InsertTCaseMemo','n_ID：未找到我方卷号-" + row + "')"); 
            }
            return 0;
        }
        #endregion 

        #region 档案位置
        private int InsertFileLocation(DataRow dataRow, int row)
        {
            string sNo = dataRow["我方卷号"].ToString().Trim();
            int numHk = GetIDbyName(sNo, 2);
            if (numHk > 0)
            {
                //当前案件是否是借出状态
                string strSql = "SELECT s_ArchiveStatus FROM dbo.TCase_Base WHERE n_CaseID=" + numHk;
                string sArchiveStatus = GetTimebySql(strSql) == null ? "" : GetTimebySql(strSql).ToString();

                if (sArchiveStatus.ToUpper().Equals("I"))
                {
                    strSql = "SELECT n_ID FROM  dbo.TCode_Employee  WHERE s_Name='" + dataRow["我方卷号"].ToString().Trim().Replace("'", "''") + "' OR s_InternalCode='" + dataRow["我方卷号"].ToString().Trim().Replace("'", "''") + "'";
                    int nUserId = GetIDbySql(strSql);
                    if (nUserId > 0)
                    {
                        strSql = "UPDATE dbo.TCase_Base SET s_ArchiveStatus='O',s_ArchivePosition=" + nUserId + " WHERE n_CaseID=" + numHk;
                        strSql += "  INSERT INTO dbo.TCase_ArchivesHistory(n_CaseID ,n_BorrowerID ,dt_BorrowerTime,s_Notes ,dt_CreateTime)"+
                                  "  VALUES  (" + numHk + "," + nUserId + ",'" + dataRow["借卷时间"].ToString() + "','借出：','"+DateTime.Now+"')";
                        return  InsertbySql(strSql);
                    }
                    else
                    {
                        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + numHk + "," + row + ",'" + sNo + "','此案件借卷人未查到！','InsertFileLocation','档案位置-" + row + "')");
                    }
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + numHk + "," + row + ",'" + sNo + "','此案件未借出状态，无法再次外借','InsertFileLocation','档案位置-" + row + "')");
                }
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + numHk + "," + row + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','InsertFileLocation','档案位置-" + row + "')");
            }
            return 0;
        }
        #endregion

        #region 更新申请人译名
        private int UpdateApplicant(DataRow dataRow, int row)
        {
            string sNo = dataRow["我方卷号"].ToString().Trim();
            int numHk = GetIDbyName(sNo, 2);
            if(numHk>0)
            { 
                int nApplicantID = GetClientandApplicantIDByName(dataRow["申请人编码"].ToString(), "Applicant");
                if(nApplicantID>0)
                {
                    string strSql = "UPDATE dbo.TCase_Applicant SET s_Name='" + dataRow["申请人译名"].ToString() + "' WHERE n_CaseID="+numHk+" AND n_ApplicantID=" + nApplicantID;
                    InsertbySql(strSql);
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + numHk + "," + row + ",'" + dataRow["申请人编码"].ToString() + "','未找到申请人为:" + dataRow["申请人编码"].ToString() + "','UpdateApplicant','更新申请人译名-" + row + "')");   
                }
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + numHk + "," + row + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','UpdateApplicant','更新申请人译名-" + row + "')");   
            }
            return 0;
        }
        #endregion

      

        #endregion


        #region  28.控制要求
        private int InsertCodeDemandType(DataRow dr)
        {
            int result = 0;
            //要求分类
            string selctSql = " select n_ID from TFCode_DemandType WHERE s_Name='" + dr["要求分类"] + "'";
            int num = GetIDbySql(selctSql);
            if (num <= 0)
            {

                string InsertSql = "INSERT INTO dbo.TFCode_DemandType (s_Name)"
                                                   + "VALUES  ( '" + dr["要求分类"] + "')";
                InsertbySql(InsertSql, 0);
                num = GetIDbySql(selctSql);
            }
            if (num > 0)
            {
                InsertCodeDemand(dr, num);
            }
            return result;
        }
        private void InsertCodeDemand(DataRow dr, int nDemandType)
        {
            string InsertSql = "INSERT INTO dbo.TCode_Demand (s_IPType , s_Type ,s_SysDemand, s_Title ,s_Description, n_IsActive , dt_CreateDate , dt_EditDate ,  n_DemandType )"
                                                 + "VALUES  ('P' ,'" + dr["要求分类"].ToString().Replace("'", "''").Replace("\r", "").Replace("\n", "") + "' ,'" + dr["系统要求代码"].ToString().Replace("'", "''").Replace("\r", "").Replace("\n", "") + "' ,'" + dr["标题"].ToString().Replace("'", "''").Replace("\r", "").Replace("\n", "") + "' ,'" + dr["描述"].ToString().Replace("'", "''").Replace("\r", "").Replace("\n", "") + "' ,1 ,GETDATE(),GETDATE(),'" + nDemandType + "')";

            string strsql = "select n_ID from TCode_Demand WHERE ";
            if (!string.IsNullOrEmpty(dr["系统要求代码"].ToString()))
            {
                strsql += "s_sysdemand='" + dr["系统要求代码"] + "' ";
            }
            else
            {
                strsql += "s_Title='" + dr["标题"].ToString().Replace("'", "''").Replace("\r", "").Replace("\n", "") + "' and n_DemandType='" + nDemandType + "'";
            }
            DataTable newTable = GetDataTablebySql(strsql);
            if (newTable.Rows.Count <= 0)
            {
                InsertbySql(InsertSql, 0);
                newTable = GetDataTablebySql(strsql);
            }
            if (newTable.Rows.Count > 0)
            {
                InsertTCodeOftenDemand(newTable.Rows[0]["n_ID"].ToString());
            }
        }
        private void InsertTCodeOftenDemand(string nDemandType)
        {
            string InsertSql = " insert into TCode_OftenDemand(n_CodeDemandID,s_CaseType,s_OftenCreator,s_OftenEditor,dt_OftenCreateDate,dt_OftenEditDate)"
                                                 + " values(" + nDemandType + ",'P','administrator','administrator','" + DateTime.Now + "','" + DateTime.Now + "')";

            DataTable newTable = GetDataTablebySql("select * from TCode_OftenDemand WHERE n_CodeDemandID=" + nDemandType + " AND s_CaseType='P' ");
            if (newTable.Rows.Count <= 0)
            {
                InsertbySql(InsertSql, 0);
            }
        }
        #endregion        

        #region 25.分案信息
        private int UpdateCaseDivisionInfo(DataRow row)
        {
            int result = 0;
            string sNo = row["我方卷号"].ToString();

            int HkNum = GetIDbyName(sNo, 2);
            if (HkNum > 0)
            {
                try
                {
                    string Sql =
                        string.Format(
                            @"UPDATE dbo.TPCase_Patent SET dt_DivSubmitDate = '{2}', b_DivisionalCaseFlag = '{1}', s_OrigAppNo = '{3}', s_OrigCaseNo = '{4}', s_DivisionAppNo = '{5}', dt_OrigAppDate = '{6}', s_DivisionCaseNo = '{7}' WHERE dbo.TPCase_Patent.n_CaseID = '{0}'",
                            HkNum,
                            "1",
                            row["分案申请提交日"] != null && !string.IsNullOrEmpty(row["分案申请提交日"].ToString()) ? row["分案申请提交日"].ToString() : string.Empty,
                            row["原案申请号"] != null && !string.IsNullOrEmpty(row["原案申请号"].ToString()) ? row["原案申请号"].ToString() : string.Empty,
                            row["原案文号"] != null && !string.IsNullOrEmpty(row["原案文号"].ToString()) ? row["原案文号"].ToString() : string.Empty,
                            row["针对的分案的申请号"] != null && !string.IsNullOrEmpty(row["针对的分案的申请号"].ToString()) ? row["针对的分案的申请号"].ToString() : string.Empty,
                            row["原案申请日"] != null && !string.IsNullOrEmpty(row["原案申请日"].ToString()) ? row["原案申请日"].ToString() : string.Empty,
                            row["针对分案文号"] != null && !string.IsNullOrEmpty(row["针对分案文号"].ToString()) ? row["针对分案文号"].ToString() : string.Empty
                            );
                    InsertbySql(Sql, 0);
                    result = 1;
                    if (!string.IsNullOrEmpty(row["原案文号"].ToString()))
                    {
                        //查找母案
                        string[] s = row["原案文号"].ToString().Split('-');
                        string strRight = s[0].ToUpper().Substring(s[0].ToUpper().Length - 3, 3);
                        int end = s[0].ToUpper().LastIndexOf("D");
                        string s2 = s[0];
                        if (strRight.Contains("D"))
                        {
                            if (end > 0)
                            {
                                s2 = s[0].Substring(0, end);
                            }
                        }
                        //查找母案关系
                        DataTable table = GetsCaseSerial(s2);
                        string Master = "";
                        if (table.Rows.Count > 0)
                        {
                            for (int i =0; i < table.Rows.Count; i++)
                            {
                                string sCaseSerial = table.Rows[i]["s_CaseSerial"].ToString();
                                string[] arry = sCaseSerial.Split('-');
                                if (!sCaseSerial.Equals(sNo))
                                {
                                    if (arry[0].Equals(s2))
                                    {
                                        Master = sCaseSerial;
                                        InsertCodeCaseRelative(HkNum, sCaseSerial, "母案");
                                    }
                                    else
                                    {
                                        InsertCodeCaseRelative(HkNum, sCaseSerial, "相关案件");
                                        if (!string.IsNullOrEmpty(Master))
                                        {
                                            InsertCodeCaseRelative(GetIDbyName(Master, 2), sCaseSerial, "分案");
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(row["针对分案文号"].ToString()))
                    {
                        InsertCodeCaseRelative(HkNum, row["针对分案文号"].ToString(), "相关案件");
                    }
                }
                catch (Exception ex)
                {
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("更新分案申请信息错误：" + ex.Message +
                                                                               "  " + sNo);
                }
            }
            else
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("不存在此我方卷号" + sNo);
            }

            return result;
        }
        //模糊查询
        private DataTable GetsCaseSerial(string sCaseSerial)
        {
            DataTable table = new DataTable();
            table.Clear();
            string sql = "SELECT s_CaseSerial FROM dbo.TCase_Base WHERE s_CaseSerial LIKE '" + sCaseSerial + "%'";
            table = GetDataTablebySql(sql);
            return table;
        }

        //增加分案信息
        private void InsertCodeCaseRelative(int HKNum, string sOrigCaseNo, string Type)
        {
            //查询母案 
            int caseID = GetIDbyName(sOrigCaseNo, 2);

            string strSql =
                   "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='分案' AND s_MasterName='母案' AND s_SlaveName='子案' AND s_IPType='P'";
            if (Type.Equals("相关案件"))
            {
                strSql =
                    "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='相关案件' AND s_MasterName='' AND s_SlaveName='' AND s_IPType='P'";
            }
            int n_ID = GetIDbySql(strSql);
            if (caseID > 0)
            {
                int NUMS =
                    GetIDbySql("SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where ((n_CaseIDA=" + HKNum +
                               " and n_CaseIDB=" + caseID + ") or  (n_CaseIDA=" + caseID + " and n_CaseIDB=" + HKNum + "))and n_CodeRelativeID=" + n_ID);
                if (NUMS <= 0 && HKNum > 0)
                {
                    if (Type.Equals("母案") || Type.Equals("相关案件"))
                    {
                        InsertTCaseCaseRelative(HKNum, caseID, n_ID, 0);
                    }
                    else
                    {
                        InsertTCaseCaseRelative(caseID, HKNum, n_ID, 0);
                    }
                }
            }
        }
        #endregion

        #region 准备客户、申请人和外代理基本信息
        #region  31.外代理信息
        #region 外代理信息-基本信息
        private int InsertTCstmrCoopAgency(DataRow _dr, int _row)
        {
            int nAgencyID = GetClientandApplicantIDByName(_dr["编码"].ToString(), "CoopAgency");
            int nClientID = GetClientandApplicantIDByName(_dr["编码"].ToString(), "Client");
            if (nClientID <= 0)
            {
                nClientID = -1;
            }
            string strSql = ""; 
            string sName = _dr["中文名称、中译名"].ToString().Trim().Replace("'", "''");//中文名
            string sNativeName = _dr["外文名称"].ToString().Trim().Replace("'", "''");//英文名
            int nLagnguage = GetLanguageIDByName(_dr["语种"].ToString().Trim().Replace("'", "''"));
            int nPayCurrency = GetCurrencyIDByName(_dr["结算币种"].ToString().Trim().Replace("'", "''"));
            string sIPType = IPtype(_dr["委托业务"].ToString().Trim().Replace("'", "''"));//P:专利；T:商标；D：域名；C：版权 O：其它
            string sCredit = _dr["信用等级"].ToString().Trim().Replace("'", "''");

            string sEmail = _dr["电子邮箱"].ToString().Trim().Replace("'", "''");
            string sFax = _dr["传真"].ToString().Trim().Replace("'", "''");
            string sWebsite = _dr["网址"].ToString().Trim().Replace("'", "''");
            string sMobile = _dr["手机"].ToString().Trim().Replace("'", "''");
            string sPhone = _dr["座机"].ToString().Trim().Replace("'", "''");
            string sNotes = _dr["备注"].ToString().Trim().Replace("'", "''");

            string sCountry = _dr["可代理国家（地区）"] != null ? "[可代理国家（地区）:" + _dr["可代理国家（地区）"].ToString().Trim().Replace("'", "''").Replace(",","|")+"]\r\n" : "";
             string billingMode = _dr["账单方式"] == null ? "" : "[账单方式:" + _dr["账单方式"].ToString() + "]\r\n";
              string dunningCycle = _dr["账期（催款周期）"] == null ? "" : "[账期（催款周期）:" + _dr["账期（催款周期）"].ToString() + "]\r\n";
              string patentPerson = _dr["专利负责人"] == null ? "" : "[专利负责人:" + _dr["专利负责人"].ToString() + "]\r\n";
            
            sNotes = sCountry + billingMode + dunningCycle + patentPerson + sNotes;
            if (nAgencyID > 0)
            {
                strSql = "update   TCstmr_CoopAgency set s_Name='" + sName + "',s_NativeName='" + sNativeName + "',s_Mobile='" + sMobile + "',s_Phone='" + sPhone + "',s_Fax='" + sFax + "',s_Website='" + sWebsite + "',s_Email='" + sEmail + "',s_Notes='" + sNotes + "'" +
                    ",n_Language=" + nLagnguage + ",n_PayCurrency=" + nPayCurrency + ",s_IPType='" + sIPType + "',s_Credit='" + sCredit + "' where n_AgencyID=" + nAgencyID;
            }
            else
            {
                strSql = "INSERT INTO dbo.TCstmr_CoopAgency(dt_CreateDate,dt_EditDate,dt_FirstCaseFromDate,dt_LastCaseFromDate,s_Creator,s_Code" +
                          ",s_Name,s_NativeName,s_Mobile,s_Phone,s_Fax,s_Website,s_Email,s_Notes,n_Language,n_PayCurrency,s_IPType,s_Credit,n_ClientID)" +
                          "VALUES  ('" + DateTime.Now + "','" + DateTime.Now + "','" + DateTime.Now + "','" + DateTime.Now + "','administrator','" + _dr["编码"].ToString() + "'" +
                               ",'" + sName + "','" + sNativeName + "','" + sMobile + "','" + sPhone + "','" + sFax + "','" + sWebsite + "','" + sEmail + "','" + sNotes + "'" +
                               "," + nLagnguage + "," + nPayCurrency + ",'" + sIPType + "','" + sCredit + "'," + nClientID + ")";
            }
            return InsertbySql(strSql);
        }

        #endregion

        #region 外代理信息-地址
        private int AddAgencyAddress(DataRow _dr, int _row)
        {
            int nAgencyID = GetClientandApplicantIDByName(_dr["编码"].ToString(), "CoopAgency");

            if (nAgencyID > 0)
            {
                AddAgencyRess(_dr, nAgencyID, "CoopAgency");
            }
            return 0;
        }
        #endregion

        #region 增加公共方法
        //币种
        private int GetCurrencyIDByName(string CurrencyName)
        {
            if (!string.IsNullOrEmpty(CurrencyName))
            {
                string strSql = "SELECT n_ID FROM dbo.TCode_Currency WHERE s_Name='" + CurrencyName + "' OR s_CurrencyCode='" + CurrencyName + "'";
                if (GetIDbySql(strSql) > 0)
                {
                    return GetIDbySql(strSql);
                }
            }
            return -1;
        }

        private void AddAgencyRess(DataRow _dr, int nAgencyID, string type)
        {
            GetIDbyClientAddress(type, nAgencyID);
            //编码	外文名称	中文名称、中译名	
            //地址1	国家1	省/州、自治区、直辖市 （中国客户）1	市县（中国客户）1	街道门牌（中国客户）1	邮编（中国客户）1	
            //地址2	国家2	省/州、自治区、直辖市 （中国客户）2	市县（中国客户）2	街道门牌2	邮编（中国客户）2	
            //地址3	国家3	省/州、自治区、直辖市 （中国客户）3	市县（中国客户）3	街道门牌（中国客户）3	邮编（中国客户）3

            int nCountry = GetAddressIDByName(_dr["国家1"].ToString().Trim().Replace("'", "''"));
            int nCountry2 = GetAddressIDByName(_dr["国家2"].ToString().Trim().Replace("'", "''"));
            int nCountry3 = GetAddressIDByName(_dr["国家3"].ToString().Trim().Replace("'", "''"));

            string sAddress = IPtype(_dr["地址1"].ToString().Trim().Replace("'", "''"));
            string sAddress2 = IPtype(_dr["地址2"].ToString().Trim().Replace("'", "''"));
            string sAddress3 = IPtype(_dr["地址3"].ToString().Trim().Replace("'", "''"));

            AddAgencyAddress(nAgencyID, nCountry, _dr["省/州、自治区、直辖市 （中国客户）1"].ToString().Trim().Replace("'", "''"), _dr["市县（中国客户）1"].ToString().Trim().Replace("'", "''"), _dr["街道门牌（中国客户）1"].ToString().Trim().Replace("'", "''"), _dr["邮编（中国客户）1"].ToString().Trim().Replace("'", "''"), sAddress);
            AddAgencyAddress(nAgencyID, nCountry2, _dr["省/州、自治区、直辖市 （中国客户）2"].ToString().Trim().Replace("'", "''"), _dr["市县（中国客户）2"].ToString().Trim().Replace("'", "''"), _dr["街道门牌2"].ToString().Trim().Replace("'", "''"), _dr["邮编（中国客户）2"].ToString().Trim().Replace("'", "''"), sAddress2);
            AddAgencyAddress(nAgencyID, nCountry3, _dr["省/州、自治区、直辖市 （中国客户）3"].ToString().Trim().Replace("'", "''"), _dr["市县（中国客户）3"].ToString().Trim().Replace("'", "''"), _dr["街道门牌（中国客户）3"].ToString().Trim().Replace("'", "''"), _dr["邮编（中国客户）3"].ToString().Trim().Replace("'", "''"), sAddress3);

        }

        private void AddAgencyAddress(int n_ClientID, int nCountry, string sState, string sCity, string s_Street, string s_ZipCode, string stype)
        {
            string strSql = "INSERT INTO dbo.TCstmr_AgencyAddress( n_AgencyID ,n_Country ,s_State ,s_City , s_Street ,s_ZipCode,s_Type)" +
                           "VALUES(" + n_ClientID + "," + nCountry + ",'" + sState + "','" + sCity + "','" + s_Street + "','" + s_ZipCode + "','" + stype + "')";

            InsertbySql(strSql);
        }

        #endregion

        #region 外代理信息-联系人
        private int AddAgencyContact(DataRow _dr, int _row)
        {
            //案件编号	email	名	语言	委托业务	手机	座机	
            //地址1类型	国家1	邮政编码1	省、州1	市县1	街道门牌1	抬头地址1	 
            //地址2类型	国家2	邮政编码2	省、州2	市县2	街道门牌2	抬头地址2	
            //地址3类型	国家3	邮政编码3	省、州3	市县3	街道门牌3	抬头地址3	

            int nAgencyID = GetClientIDByCaseSerialandCode(_dr["案件编号"].ToString(), "Agency", _dr["外代理"].ToString());
            int nCaseID = GetIDbyName(_dr["案件编号"].ToString().Trim(), 2);
            string strSql = "";
            if (nAgencyID > 0)
            {
                int nLanguage = GetLanguageIDByName(_dr["语言"].ToString().Trim().Replace("'", "''"));
                string sPhone = _dr["座机"].ToString().Trim().Replace("'", "''");
                string sMobile = _dr["手机"].ToString().Trim().Replace("'", "''");
                string sFirstName = _dr["名"].ToString().Trim().Replace("'", "''");
                string sIPType = _dr["委托业务"].ToString().Trim().Replace("'", "''");
                string sEmail = _dr["email"].ToString().Trim().Replace("'", "''");
                int nContactID = GetClientIDByCaseSerial(sFirstName, sEmail, nAgencyID, "Agency");
                if (nContactID > 0)
                {
                    strSql = "update TCstmr_AgencyContact set s_Phone='" + sPhone + "',s_Mobile='" + sMobile + "',s_IPType='" + sIPType + "',n_Language=" + nLanguage + " where n_ContactID=" + nContactID;
                }
                else
                {
                    strSql = " INSERT INTO dbo.TCstmr_AgencyContact( n_AgencyID ,s_FirstName , s_IPType ,n_Language ,s_Phone ,s_Mobile,s_Email)" +
                        " VALUES  ( " + nAgencyID + ",'" + sFirstName + "','" + sIPType + "'," + nLanguage + ",'" + sPhone + "','" + sMobile + "','" + sEmail + "')";

                }

                if (InsertbySql(strSql) > 0)
                {
                    nContactID = GetClientIDByCaseSerial(sFirstName, sEmail, nAgencyID, "Agency");
                    AddRess(_dr, nContactID, "Agency");
                    if (nCaseID > 0)
                    {
                        InsertTCaseContact(nCaseID, "Agency", nContactID, "");
                    }
                }
            }
            else
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(_row + "未找到案件:" + _dr["案件编号"].ToString() + _dr["名"].ToString() + _dr["email"].ToString());
            }
            return 0;
        }
        #endregion

        #region 外代理信息-财务
        private int AddAgencyBill(DataRow _dr)
        {
            //编码	外文名称	中文名称、中译名	收款银行名称	收款银行地址	
            //收款人名称	收款人地址	收款人账户号码	IBAN	Swift Code	ABA 
            int nAgencyID = GetClientandApplicantIDByName(_dr["编码"].ToString(), "CoopAgency");
            string strSql = "";
            if (nAgencyID > 0)
            {
                strSql = "update TCstmr_CoopAgency set s_BeneficiaryBankName='" + _dr["收款银行名称"].ToString() + "',s_BeneficiaryBankAddress='" + _dr["收款银行地址"].ToString() + "',s_BeneficiaryName='" + _dr["收款人名称"].ToString() + "'" +
                         ",s_BeneficiaryAddress='" + _dr["收款人地址"].ToString() + "',s_BeneficiaryAccountNumber='" + _dr["收款人账户号码"].ToString() + "',s_IBAN='" + _dr["IBAN"].ToString() + "',s_SwiftCode='" + _dr["Swift Code"].ToString() + "'" +
                    ",S_ABA='" + _dr["ABA"].ToString() + "'   where n_AgencyID =" + nAgencyID;
                return InsertbySql(strSql);
            }
            return 0;
        }
        #endregion

        #endregion

        #region  30.申请人信息
        #region 申请人信息-基本信息
        private int InsertTCstmrApplicant(DataRow _dr, int _row)
        {
            int nApplicantID = GetClientandApplicantIDByName(_dr["申请人代码"].ToString(), "Applicant");
            int nClientID = GetClientandApplicantIDByName(_dr["申请人代码"].ToString(), "Client");
            if (nClientID <= 0)
            {
                nClientID = -1;
            }

            string sName = _dr["中文名称、中译名"].ToString().Trim().Replace("'", "''");//中文名
            string sNativeName = _dr["外文名称"].ToString().Trim().Replace("'", "''");//英文名
            string sEmail = _dr["电子邮箱"].ToString().Trim().Replace("'", "''");
            string sFax = _dr["传真"].ToString().Trim().Replace("'", "''");
            string sWebsite = _dr["网址"].ToString().Trim().Replace("'", "''");
            string sMobile = _dr["手机"].ToString().Trim().Replace("'", "''");
            string sPhone = _dr["座机"].ToString().Trim().Replace("'", "''");
            string s_TrustDeedNo = _dr["总委托书"].ToString().Trim().Replace("'", "''");
            string sPayFeePerson = _dr["电子申请缴费人"].ToString().Trim().Replace("'", "''");
            string sAppType = _dr["申请人类型"].ToString().Trim().Replace("'", "''");
            string sIDNumber = _dr["身份证号或组织机构代码或统一社会信用代码"].ToString().Trim().Replace("'", "''");
            //费减资格备案年度	费减备案编号 
            string sFeeMitigationYear = _dr["费减资格备案年度"].ToString().Trim().Replace("'", "''");
            string sFeeMitigationNum = _dr["费减备案编号"].ToString().Trim().Replace("'", "''");
            //转换后数据
            int nCountry = GetAddressIDByName(_dr["国籍、注册国家（地区）"].ToString().Trim().Replace("'", "''"));


            string qualityRating = _dr["信用等级"] == null ? "" : "[信用等级:" + _dr["信用等级"].ToString() + "]\r\n";
            string billingMode = _dr["账单方式"] == null ? "" : "[账单方式:" + _dr["账单方式"].ToString() + "]\r\n";
            string currency = _dr["币种"] == null ? "" : "[币种:" + _dr["币种"].ToString() + "]";
            string dunningCycle = _dr["账期（催款周期）"] == null ? "" : "[账期（催款周期）:" + _dr["账期（催款周期）"].ToString() + "]\r\n";
            string language = _dr["通信语种"] == null ? "" : "[通信语种:" + _dr["通信语种"].ToString().Replace(",","|") + "]\r\n";
            string patentPerson = _dr["专利负责人"] == null ? "" : "[专利负责人:" + _dr["专利负责人"].ToString() + "]\r\n";
            string sNotes = _dr["备注"].ToString().Trim().Replace("'", "''");

            string ctx = qualityRating + billingMode + currency + dunningCycle + language + patentPerson + sNotes;

            if (!string.IsNullOrEmpty(sName) || !string.IsNullOrEmpty(sNativeName))
            {
                string strSql = "";
                if (nApplicantID > 0)
                {
                    strSql = "update    TCstmr_Applicant set s_Name='" + sName + "',s_NativeName='" + sNativeName + "',s_Mobile='" + sMobile + "',s_Phone='" + sPhone + "',s_Fax='" + sFax + "',s_Website='" + sWebsite + "',s_Email='" + sEmail + "',s_Notes='" + ctx.Replace("'","''") + "'" +
                        ",s_AppType='" + sAppType + "',s_IDNumber='" + sIDNumber + "',n_Country=" + nCountry + ",n_ClientID=" + nClientID + ",s_FeeMitigationYear='" + sFeeMitigationYear + "',s_FeeMitigationNum='" + sFeeMitigationNum + "'";
                    if (sPayFeePerson.Equals("是") || sPayFeePerson.Equals("Y"))
                    {
                        strSql += ",s_PayFeePerson='Y'";
                    }
                    else
                    {
                        strSql += ",s_PayFeePerson='N'";
                    }
                    if (!string.IsNullOrEmpty(s_TrustDeedNo))
                    {
                        strSql += ",s_TrustDeedNo='" + s_TrustDeedNo + "',s_HasTrustDeed='Y',dt_EditDate='" + DateTime.Now + "'";
                    }
                    strSql += " where n_AppID=" + nApplicantID;
                }
                else
                { 
                    strSql = "INSERT INTO dbo.TCstmr_Applicant(dt_CreateDate,dt_EditDate,dt_FirstCaseFromDate,dt_LastCaseFromDate,s_Creator,s_AppCode" +
                              ",s_Name,s_NativeName,s_Mobile,s_Phone,s_Fax,s_Website,s_Email,s_Notes,s_AppType,s_IDNumber,n_Country,n_ClientID,s_FeeMitigationYear,s_FeeMitigationNum";
                    string Values = "VALUES  ('" + DateTime.Now + "','" + DateTime.Now + "','" + DateTime.Now + "','" + DateTime.Now + "','administrator','" + _dr["申请人代码"].ToString() + "'" +
                                   ",'" + sName + "','" + sNativeName + "','" + sMobile + "','" + sPhone + "','" + sFax + "','" + sWebsite + "','" + sEmail + "','" + ctx.Replace("'", "''") + "','" + sAppType + "','" + sIDNumber + "'," + nCountry + "," + nClientID +
                                   ",'" + sFeeMitigationYear + "','" + sFeeMitigationNum + "'";

                    if (sPayFeePerson.Equals("是") || sPayFeePerson.Equals("Y"))
                    {
                        strSql += ",s_PayFeePerson";
                        Values += ",'Y'";
                    }
                    else
                    {
                        strSql += ",s_PayFeePerson";
                        Values += ",'N'";
                    }
                    if (!string.IsNullOrEmpty(s_TrustDeedNo))
                    {
                        strSql += ",s_TrustDeedNo,s_HasTrustDeed";
                        Values += ",'" + s_TrustDeedNo + "','Y'";
                    }
                    strSql = strSql + ")" + Values + ")";
                }
                return InsertbySql(strSql);
            }
            return 1;
        }
        #endregion

        #region 申请人-地址
        private int AddAppAddress(DataRow _dr, int _row)
        {
            int nApplicantID = GetClientandApplicantIDByName(_dr["申请人代码"].ToString(), "Applicant");

            if (nApplicantID > 0)
            {
                return AddAPPRess(_dr, nApplicantID, "Applicant");
            }
            return 0;
        }
        #endregion

        #region 增加公共方法
        //地址
        private int AddAPPRess(DataRow _dr, int nAppID, string type)
        {
            if (type.Equals("Applicant"))
            {
                string strSql = "delete TCstmr_AppAddress WHERE n_AppID=" + nAppID;
                GetIDbySql(strSql);
            }
            int nCountry = GetAddressIDByName(_dr["居所或营业场所 （页签2：国家）"].ToString().Trim().Replace("'", "''"));
            int nCountry2 = GetAddressIDByName(_dr["国家2"].ToString().Trim().Replace("'", "''"));
            int nCountry3 = GetAddressIDByName(_dr["国家3"].ToString().Trim().Replace("'", "''"));

            string sAddress = IPtype(_dr["地址1"].ToString().Trim().Replace("'", "''"));
            string sAddress2 = IPtype(_dr["地址2"].ToString().Trim().Replace("'", "''"));
            string sAddress3 = IPtype(_dr["地址3"].ToString().Trim().Replace("'", "''"));

            //申请人代码	
            //地址1	居所或营业场所 （页签2：国家）	省/州、自治区、直辖市 （中国客户）	市县（中国客户）	街道门牌（中国客户）	邮编（中国客户）	
            //地址2	国家	省/州、自治区、直辖市 （中国客户）	市县（中国客户）	街道门牌（中国客户）	邮编（中国客户）	
            //地址3	国家	省/州、自治区、直辖市 （中国客户）	市县（中国客户）	街道门牌（中国客户）	邮编（中国客户）
            AddAPPAddress(nAppID, nCountry, _dr["省/州、自治区、直辖市 （中国客户）1"].ToString().Trim().Replace("'", "''"), _dr["市县（中国客户）1"].ToString().Trim().Replace("'", "''"), _dr["街道门牌（中国客户）1"].ToString().Trim().Replace("'", "''"), _dr["邮编（中国客户）1"].ToString().Trim().Replace("'", "''"), sAddress);
            AddAPPAddress(nAppID, nCountry2, _dr["省/州、自治区、直辖市 （中国客户）2"].ToString().Trim().Replace("'", "''"), _dr["市县（中国客户）2"].ToString().Trim().Replace("'", "''"), _dr["街道门牌（中国客户）2"].ToString().Trim().Replace("'", "''"), _dr["邮编（中国客户）2"].ToString().Trim().Replace("'", "''"), sAddress2);
            AddAPPAddress(nAppID, nCountry3, _dr["省/州、自治区、直辖市 （中国客户）3"].ToString().Trim().Replace("'", "''"), _dr["市县（中国客户）3"].ToString().Trim().Replace("'", "''"), _dr["街道门牌（中国客户）3"].ToString().Trim().Replace("'", "''"), _dr["邮编（中国客户）3"].ToString().Trim().Replace("'", "''"), sAddress3);

            return 1;
        }
        private void AddAPPAddress(int n_AppID, int nCountry, string sState, string sCity, string s_Street, string s_ZipCode, string sType)
        {
            string strSql = "INSERT INTO dbo.TCstmr_AppAddress( n_AppID ,n_Country ,s_State ,s_City , s_Street ,s_ZipCode,s_Type)" +
                           "VALUES(" + n_AppID + "," + nCountry + ",'" + sState + "','" + sCity + "','" + s_Street + "','" + s_ZipCode + "','" + sType + "')";

            InsertbySql(strSql);
        }
        #endregion

        #region 申请人-联系人
        private int AddAppContact(DataRow _dr, int _row)
        {
            // 	email	名	语言	委托业务	手机	座机	
            //地址1类型	国家	邮政编码	省、州	市县	街道门牌	抬头地址	
            //地址2类型	国家	邮政编码	省、州	市县	街道门牌	抬头地址
            //地址3类型	国家	邮政编码	省、州	市县	街道门牌	抬头地址
            //if (_dr["案件编号"].ToString() == "16ZX1101-1836-ZYA")
            //{
                if (!string.IsNullOrEmpty(_dr["申请人代码"].ToString()))
                {
                    int nAppID = GetAppIDBysAppCode(_dr["申请人代码"].ToString());
                    int nCaseID = GetIDbyName(_dr["案件编号"].ToString().Trim(), 2);

                    if (nAppID > 0)
                    {
                        int nLanguage = GetLanguageIDByName(_dr["语言"].ToString().Trim().Replace("'", "''"));
                        string sPhone = _dr["座机"].ToString().Trim().Replace("'", "''");
                        string sMobile = _dr["手机"].ToString().Trim().Replace("'", "''");
                        string sFirstName = _dr["名"].ToString().Trim().Replace("'", "''");
                        string sIPType = _dr["委托业务"].ToString().Trim().Replace("'", "''");
                        string sEmail = _dr["email"].ToString().Trim().Replace("'", "''");
                        string sDepartment = _dr["部门"].ToString().Trim().Replace("'", "''");
                        string sJobTitle = _dr["职位"].ToString().Trim().Replace("'", "''");
                        int nContactID = GetContactIDBysAppCode(sFirstName, sEmail, _dr["申请人代码"].ToString());
                        string strSql = "";
                        if (nContactID > 0)
                        {
                            strSql = "update TCstmr_AppContact set s_Phone='" + sPhone + "',s_Mobile='" + sMobile + "',s_IPType='" + sIPType + "',n_Language=" + nLanguage + ",s_Department='" + sDepartment + "',s_JobTitle='" + sJobTitle + "' where n_AppID=" + nAppID;
                        }
                        else
                        {
                            strSql = " INSERT INTO dbo.TCstmr_AppContact( n_AppID ,s_FirstName , s_IPType ,n_Language ,s_Phone ,s_Mobile,s_Email,s_Department,s_JobTitle)" +
                                " VALUES  ( " + nAppID + ",'" + sFirstName + "','" + sIPType + "'," + nLanguage + ",'" + sPhone + "','" + sMobile + "','" + sEmail + "','" + sDepartment + "','" + sJobTitle + "')";
                        }
                        if (InsertbySql(strSql) > 0)
                        {
                            nContactID = GetContactIDBysAppCode(sFirstName, sEmail, _dr["申请人代码"].ToString());
                            AddRess(_dr, nContactID, "AppContact");
                            if (nCaseID > 0)
                            {
                                InsertTCaseContact(nCaseID, "Applicant", nContactID, "");
                            }
                        }
                       
                    }
                    else
                    {
                        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(0," + _row + ",'" + _dr["案件编号"].ToString() + "','未找到申请人为:" + _dr["申请人代码"].ToString() + _dr["名"].ToString() + _dr["email"].ToString() + "','TCaseLaw','申请人-联系人-" + _row + "')");
                    }
                }
            //}
            return 1;
        }
        private int GetAppIDBysAppCode(string sAppCode)
        {
            string strSql = "SELECT n_AppID FROM TCstmr_Applicant WHERE s_AppCode='" + sAppCode + "'";
            return GetIDbySql(strSql);
        }
        ////查找联系人是否存在
        private int GetContactIDBysAppCode(string sName, string sEmail, string sAppCode)
        {
            string strSql = "SELECT n_ContactID FROM dbo.TCstmr_AppContact WHERE n_AppID in (SELECT n_AppID FROM TCstmr_Applicant WHERE s_AppCode='" + sAppCode + "') AND s_FirstName='" + sName.Trim() + "' AND s_Email='" + sEmail.Trim() + "'";
            return GetIDbySql(strSql);
        }
        #endregion

        #region 申请人-财务
        private int AddAppBill(DataRow _dr)
        {
            //申请人代码	单位名称	纳税人识别号	注册地址	注册电话	开户行名称	银行账号	发票抬头
            int nAppID = GetClientandApplicantIDByName(_dr["申请人代码"].ToString(), "Applicant");
            string strSql = "";
            if (nAppID > 0)
            {
                strSql = "update TCstmr_Applicant set s_AccountName='" + _dr["单位名称"].ToString() + "',s_TaxCode='" + _dr["纳税人识别号"].ToString() + "',s_RegAddress='" + _dr["注册地址"].ToString() + "',s_RegTel='" + _dr["注册电话"].ToString() + "',s_BankName='" + _dr["开户行名称"].ToString() + "',s_AccountNo='" + _dr["银行账号"].ToString() + "',s_InvoiceTo='" + _dr["发票抬头"].ToString() + "'" +
                    "  where n_AppID =" + nAppID;
                return InsertbySql(strSql);
            }
            return 0;
        }

        #endregion

        #endregion

        #region  29.客户信息
        #region 客户信息-基本信息
        private int InsertTCstmrClient(DataRow _dr, int _row)
        {
            int nClientID = GetClientandApplicantIDByName(_dr["代码"].ToString(), "Client");
            string strSql = "";

            string sName = _dr["中文名称"].ToString().Trim().Replace("'", "''");//中文名
            string sNativeName = _dr["英文名称"].ToString().Trim().Replace("'", "''");//英文名
            string sEmail = _dr["电子邮箱"].ToString().Trim().Replace("'", "''");
            string sFax = _dr["传真"].ToString().Trim().Replace("'", "''");
            string sWebsite = _dr["网址"].ToString().Trim().Replace("'", "''");
            string sMobile = _dr["手机"].ToString().Trim().Replace("'", "''");
            string sPhone = _dr["座机"].ToString().Trim().Replace("'", "''");
            string sNotes = _dr["备注"].ToString().Trim().Replace("'", "''");
            string sType = _dr["客户类别"].ToString().Trim().Replace("'", "''");
            string sCredit = _dr["信用等级"].ToString().Trim().Replace("'", "''");
            string sState = _dr["省、州"].ToString().Trim().Replace("'", "''");
            string sCity = _dr["市县"].ToString().Trim().Replace("'", "''");

            //转换后数据
            int nCountry = GetAddressIDByName(_dr["国家"].ToString().Trim().Replace("'", "''"));
            string sIPType = IPtype(_dr["委托业务"].ToString().Trim().Replace("'", "''"));//P:专利；T:商标；D：域名；C：版权 O：其它
            int nPatentChargerID = GetEmployeeIDByName(_dr["专利负责人"].ToString().Trim().Replace("'", "''"));

             string[] d = _dr["通信语种"].ToString().Trim().Replace("'", "''").Split(',');
            int nLanguage = 0;
            if(d.Length>0 && !string.IsNullOrEmpty(d[0]))
            {
                nLanguage = GetLanguageIDByName(d[0]);
            }
            string dunningCycle = _dr["账期（催款周期）"] == null ? "" : "[账期（催款周期）:" + _dr["账期（催款周期）"].ToString() + "]\r\n";
            string slanguage = _dr["通信语种"] == null ? "" : "[通信语种:" + _dr["通信语种"].ToString().Replace(",", "|") + "]\r\n";

            sNotes = dunningCycle + slanguage + sNotes; 
           
            if (nClientID > 0)
            {
                strSql = "update TCstmr_Client set s_Name='" + sName + "',s_NativeName='" + sNativeName + "',s_Email='" + sEmail + "',s_Fax='" + sFax + "',s_Website='" + sWebsite + "',s_Mobile='" + sMobile + "',s_Phone='" + sPhone + "',s_Notes='" + sNotes + "',s_Type='" + sType + "',s_Credit='" + sCredit + "'" +
                         ",n_Country=" + nCountry + ",s_State='" + sState + "',s_City='" + sCity + "',n_Language=" + nLanguage + ",n_PatentChargerID=" + nPatentChargerID + ",s_IPType='" + sIPType + "',dt_EditDate='" + DateTime.Now + "' where n_ClientID=" + nClientID + "";
                InsertClientFeePolicy(_dr["币种"].ToString(), nClientID, _row);
                return InsertbySql(strSql);
            }
            else
            {
                strSql =
                              "INSERT INTO dbo.TCstmr_Client(dt_CreateDate,dt_EditDate,s_Creator,s_ClientCode,s_Name,s_NativeName,s_Email,s_Fax,s_Website,s_Mobile,s_Phone,s_Notes,s_Type,s_Credit,n_Country,s_State,s_City,n_Language,n_PatentChargerID,s_IPType) " +
                              "VALUES('" + DateTime.Now + "','" + DateTime.Now + "','administrator','" + _dr["代码"].ToString() + "','" + sName + "','" + sNativeName + "','" + sEmail + "','" + sFax + "','" + sWebsite + "','" + sMobile + "','" + sPhone + "','" + sNotes + "','" + sType + "','" + sCredit + "'" +
                               "," + nCountry + ",'" + sState + "','" + sCity + "'," + nLanguage + "," + nPatentChargerID + ",'" + sIPType + "')";
                if(InsertbySql(strSql)>0)
                {
                    nClientID = GetClientandApplicantIDByName(_dr["代码"].ToString(), "Client");
                    InsertClientFeePolicy(_dr["币种"].ToString(), nClientID, _row);
                }
                return 1;
            } 
        }

        private void InsertClientFeePolicy(string currency, int nClientID,int row)
        {
            string strSql = "";
            string[] arr = currency.Split(',');
            foreach (var item in arr)
            {
                string selectSql = "SELECT n_ID FROM dbo.TCode_Currency WHERE s_Name='" + item.Trim() + "' OR s_CurrencyCode='" + item.Trim() + "'";
                int current = GetIDbySql(selectSql);
                if (current > 0)
                {
                    strSql =
                        "INSERT INTO dbo.TCstmr_ClientFeePolicy( n_ClientID , n_ChargeCurrency , s_IPType ,dt_EditDate , s_BusinessType ,s_PTCType)" +
                        "VALUES (" + nClientID + "," + current + ",'P','" + DateTime.Now + "','-1','-1')   ";
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + nClientID + "," + row + ",'" + nClientID + "','客户ID:" + nClientID + "','InsertClientFeePolicy','户信息-基本信息-" + row + "')"); 
                }
            }
            InsertbySql(strSql);
        }

        //客户ID
        private int GetClientandApplicantIDByName(string sCode, string type)
        {
            string strSql = "";
            if (!string.IsNullOrEmpty(sCode) && type.Equals("Client"))
            {
                strSql = "select  n_ClientID from TCstmr_Client where s_ClientCode='" + sCode.Trim() + "'";
                return GetIDbySql(strSql);
            }
            else if (!string.IsNullOrEmpty(sCode) && type.Equals("Applicant"))
            {
                strSql = "select  n_AppID from TCstmr_Applicant where s_AppCode='" + sCode.Trim() + "'";
                return GetIDbySql(strSql);
            }
            else if (!string.IsNullOrEmpty(sCode) && type.Equals("CoopAgency"))
            {
                strSql = "select  n_AgencyID from TCstmr_CoopAgency WHERE s_Code='" + sCode.Trim() + "'";
                return GetIDbySql(strSql);
            }
            return 0;
        }


        //查找地址信息
        private int GetAddressIDByName(string countryName)
        {
            if (!string.IsNullOrEmpty(countryName))
            {
                string strSql = "SELECT n_ID FROM dbo.TCode_Country  WHERE  s_Name='" + countryName.Trim() + "' OR s_CountryCode='" + countryName.Trim() + "' OR s_OtherName='" + countryName.Trim() + "'";
                if (GetIDbySql(strSql) > 0)
                {
                    return GetIDbySql(strSql);
                }
            }
            return -1;
        }

        //查找语言ID
        private int GetLanguageIDByName(string sLanguageName)
        { 
            if (!string.IsNullOrEmpty(sLanguageName))
            {
                string strSql = "SELECT n_ID FROM dbo.TCode_Language  WHERE  s_Name='" + sLanguageName.Trim() + "' or s_LanguageCode='" + sLanguageName.Trim() + "' or s_OtherName='" + sLanguageName.Trim() + "'";
                return GetIDbySql(strSql);
            }
            return 0;
        }

        //专利负责人
        private int GetEmployeeIDByName(string sEmployeeName)
        { 
            if (!string.IsNullOrEmpty(sEmployeeName))
            {
                string strSql = "SELECT n_ID FROM dbo.TCode_Employee  WHERE  s_Name='" + sEmployeeName.Trim() + "' or s_InternalCode='" + sEmployeeName.Trim() + "'";
                return GetIDbySql(strSql);
            }
            return 0;
        }
        #endregion

        #region 客户-地址
        private int AddClientAddress(DataRow _dr, int _row)
        {
            int nClientID = GetClientandApplicantIDByName(_dr["代码"].ToString(), "Client");
            if (nClientID > 0)
            {
                return AddRess(_dr, nClientID, "Client");
            }
            return 0;
        }
        #endregion

        #region 增加公共方法
        //是否存在此地址
        private void GetIDbyClientAddress(string type, int nContactID)
        {
            string strSql = "delete TCstmr_ClientAddress WHERE n_ClientID=" + nContactID;
            if (type.Equals("Contact"))
            {
                strSql = "delete TCstmr_ClientConAddress WHERE n_ContactID=" + nContactID;
            }
            else if (type.Equals("AppContact"))
            {
                strSql = "delete TCstmr_AppConAddress WHERE n_ContactID=" + nContactID;
            }
            else if (type.Equals("CoopAgency"))
            {
                strSql = "delete TCstmr_AgencyAddress WHERE n_AgencyID=" + nContactID;
            }
            else if (type.Equals("Agency"))
            {
                strSql = "delete TCstmr_AgencyConAddress WHERE n_ContactID=" + nContactID;
            }
            GetIDbySql(strSql);
        }

        private void AddClientAddress(int n_ClientID, int nCountry, string sState, string sCity, string s_Street, string s_ZipCode, string type, string saddress, int nSequence, string sTitleAddress)
        {
            sState = sState.Replace(",", ",,");
            sCity = sCity.Replace(",", ",,");
            s_Street = s_Street.Replace(",", ",,");
            string strSql = "INSERT INTO dbo.TCstmr_ClientAddress( n_ClientID ,n_Country ,s_State ,s_City , s_Street ,s_ZipCode,s_Type,s_TitleAddress)" +
                           "VALUES(" + n_ClientID + "," + nCountry + ",'" + sState + "','" + sCity + "','" + s_Street + "','" + s_ZipCode + "','" + saddress + "','" + sTitleAddress + "')";
            if (type.Equals("Contact"))
            {
                strSql = "INSERT INTO dbo.TCstmr_ClientConAddress( n_ContactID ,n_Country ,s_State ,s_City , s_Street ,s_ZipCode,s_Type,s_TitleAddress)" +
                           "VALUES(" + n_ClientID + "," + nCountry + ",'" + sState + "','" + sCity + "','" + s_Street + "','" + s_ZipCode + "','" + saddress + "','" + sTitleAddress + "')";
            }
            else if (type.Equals("AppContact"))
            {
                strSql = "INSERT INTO dbo.TCstmr_AppConAddress( n_ContactID ,n_Country ,s_State ,s_City , s_Street ,s_ZipCode,s_Type,n_Sequence,s_TitleAddress)" +
                               "VALUES(" + n_ClientID + "," + nCountry + ",'" + sState + "','" + sCity + "','" + s_Street + "','" + s_ZipCode + "','" + saddress + "'," + nSequence + ",'" + sTitleAddress + "')";
            }
            else if (type.Equals("Agency"))
            {
                strSql = "INSERT INTO dbo.TCstmr_AgencyConAddress( n_ContactID ,n_Country ,s_State ,s_City , s_Street ,s_ZipCode,s_Type)" +
                            "VALUES(" + n_ClientID + "," + nCountry + ",'" + sState + "','" + sCity + "','" + s_Street + "','" + s_ZipCode + "','" + saddress + "')";
            }
            InsertbySql(strSql);
        }

        //地址
        private int AddRess(DataRow _dr, int nClientID, string type)
        {
            GetIDbyClientAddress(type, nClientID);
            int nCountry = GetAddressIDByName(_dr["国家1"].ToString().Trim().Replace("'", "''"));
            int nCountry2 = GetAddressIDByName(_dr["国家2"].ToString().Trim().Replace("'", "''"));
            int nCountry3 = GetAddressIDByName(_dr["国家3"].ToString().Trim().Replace("'", "''"));

            string sAddress = IPtype(_dr["地址1类型"].ToString().Trim().Replace("'", "''"));
            string sAddress2 = IPtype(_dr["地址2类型"].ToString().Trim().Replace("'", "''"));
            string sAddress3 = IPtype(_dr["地址3类型"].ToString().Trim().Replace("'", "''"));

            AddClientAddress(nClientID, nCountry, _dr["省、州1"].ToString().Trim().Replace("'", "''"), _dr["市县1"].ToString().Trim().Replace("'", "''"), _dr["街道门牌1"].ToString().Trim().Replace("'", "''"), _dr["邮政编码1"].ToString().Trim().Replace("'", "''"), type, sAddress, 0, _dr["抬头地址1"].ToString().Trim().Replace("'", "''"));
            AddClientAddress(nClientID, nCountry2, _dr["省、州2"].ToString().Trim().Replace("'", "''"), _dr["市县2"].ToString().Trim().Replace("'", "''"), _dr["街道门牌2"].ToString().Trim().Replace("'", "''"), _dr["邮政编码2"].ToString().Trim().Replace("'", "''"), type, sAddress2, 1, _dr["抬头地址2"].ToString().Trim().Replace("'", "''"));
            AddClientAddress(nClientID, nCountry3, _dr["省、州3"].ToString().Trim().Replace("'", "''"), _dr["市县3"].ToString().Trim().Replace("'", "''"), _dr["街道门牌3"].ToString().Trim().Replace("'", "''"), _dr["邮政编码3"].ToString().Trim().Replace("'", "''"), type, sAddress3, 2, _dr["抬头地址3"].ToString().Trim().Replace("'", "''"));

            return 1;
        }

        //类型转换
        private string IPtype(string _content)
        {
            //P:专利；T:商标；D：域名；C：版权 O：其它B, M, O, E
            string strHtml = "";
            string[] arry = _content.Replace('，', ',').Split(',');
            for (int i = 0; i < arry.Length; i++)
            {
                if (!string.IsNullOrEmpty(arry[i]))
                {
                    if (arry[i].Equals("专利"))
                    {
                        strHtml += "P,";
                    }
                    else if (arry[i].Equals("商标"))
                    {
                        strHtml += "T,";
                    }
                    else if (arry[i].Equals("域名"))
                    {
                        strHtml += "D,";
                    }
                    else if (arry[i].Equals("版权"))
                    {
                        strHtml += "C,";
                    }
                    else if (arry[i].Equals("法律&其他案"))
                    {
                        strHtml += "O,";
                    }//账单 B  转函地址 M 办公地址 O  办公地址（外文）E   
                    else if (arry[i].Equals("账单地址"))
                    {
                        strHtml += "B,";
                    }
                    else if (arry[i].Equals("转函地址"))
                    {
                        strHtml += "M,";
                    }
                    else if (arry[i].Equals("办公地址"))
                    {
                        strHtml += "O,";
                    }
                    else if (arry[i].Equals("办公地址(外文)"))
                    {
                        strHtml += "E,";
                    }
                }
            }
            if (strHtml.Length > 0)
            {
                strHtml = strHtml.Substring(0, strHtml.Length - 1);
            }
            return strHtml;
        }
        #endregion

        #region 客户-联系人
        private int AddClientContact(DataRow _dr, int _row)
        {
            int nClientID = GetClientIDByCaseSerialandCode(_dr["案件编号"].ToString(), "Client", _dr["客户代码"].ToString());
            int nCaseID = GetIDbyName(_dr["案件编号"].ToString().Trim(), 2);
            string strSql = "";
            if (nClientID > 0)
            {
                int nLanguage = GetLanguageIDByName(_dr["语言"].ToString().Trim().Replace("'", "''"));
                string sPhone = _dr["座机"].ToString().Trim().Replace("'", "''");
                string sMobile = _dr["手机"].ToString().Trim().Replace("'", "''");
                string sFirstName = _dr["名"].ToString().Trim().Replace("'", "''");
                string sIPType = _dr["委托业务"].ToString().Trim().Replace("'", "''");
                string sEmail = _dr["email"].ToString().Trim().Replace("'", "''");
                int nContactID = GetClientIDByCaseSerial(sFirstName, sEmail, nClientID, "Client");
                if (nContactID > 0)
                {
                    strSql = "update TCstmr_ClientContact set s_Phone='" + sPhone + "',s_Mobile='" + sMobile + "',s_IPType='" + sIPType + "',n_Language=" + nLanguage + " where n_ContactID=" + nContactID;
                }
                else
                {
                    strSql = " INSERT INTO dbo.TCstmr_ClientContact( n_ClientID ,s_FirstName , s_IPType ,n_Language ,s_Phone ,s_Mobile,s_Email)" +
                        " VALUES  ( " + nClientID + ",'" + sFirstName + "','" + sIPType + "'," + nLanguage + ",'" + sPhone + "','" + sMobile + "','" + sEmail + "')";

                }
                if (InsertbySql(strSql) > 0)
                {
                    nContactID = GetClientIDByCaseSerial(sFirstName, sEmail, nClientID, "Client");
                    AddRess(_dr, nContactID, "Contact");
                    if (nCaseID > 0)
                    {
                        InsertTCaseContact(nCaseID, "Client", nContactID, "");
                    }
                }
                return 1;
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(0," + _row + ",'" + _dr["案件编号"].ToString() + "','未找到案件为:" + _dr["案件编号"].ToString() + _dr["名"].ToString() + _dr["email"].ToString() + "','TCaseLaw','客户-联系人-" + _row + "')");
            }
            //信函抬头	账单抬头	   
            return 0;
        }
        //增加联系人与案件关系
        private void InsertTCaseContact(int nCaseID, string sContactType, int nContactID, string sIdentity)
        {
            string strSql = "SELECT TOP 1 n_Sequence FROM  dbo.TCase_Contact WHERE n_CaseID=" + nCaseID + " ORDER BY n_Sequence DESC ";
            int nSequence = GetIDbySql(strSql) + 1;
            strSql = "SELECT COUNT(*) AS sumcount FROM  dbo.TCase_Contact WHERE n_CaseID=" + nCaseID + " AND s_ContactType='" + sContactType + "' AND n_ContactID=" + nContactID;
            if (GetIDbySql(strSql) <= 0)
            {
                strSql = "INSERT INTO dbo.TCase_Contact( n_CaseID ,s_ContactType , n_ContactID ,n_Sequence ,s_Identity)" +
                        "VALUES  (" + nCaseID + ",'" + sContactType + "'," + nContactID + "," + nSequence + ",'" + sIdentity + "')";
                InsertbySql(strSql);
            }
        }


        //根据案件查询客户ID
        private int GetClientIDByCaseSerialandCode(string sCaseSerial, string type, string sClientCode)
        {
            string strSql = " SELECT a.n_ClientID FROM dbo.TCase_Base  a   LEFT JOIN dbo.TCstmr_Client b ON a.n_ClientID = b.n_ClientID  WHERE s_CaseSerial ='" + sCaseSerial + "' AND s_ClientCode='" + sClientCode + "'";
            if (type.Equals("Agency"))
            {
                strSql = "  SELECT n_CoopAgencyToID FROM dbo.TCase_Base  a   LEFT JOIN dbo.TCstmr_CoopAgency b ON a.n_CoopAgencyToID = b.n_AgencyID    WHERE s_CaseSerial='" + sCaseSerial + "' AND s_Code='" + sClientCode + "'";
            }
            return GetIDbySql(strSql);
        }
        //查找联系人是否存在
        private int GetClientIDByCaseSerial(string sName, string sEmail, int nClientID, string type)
        {
            string strSql = "  SELECT n_ContactID FROM dbo.TCstmr_ClientContact WHERE n_ClientID=" + nClientID + " AND s_FirstName='" + sName.Trim() + "' AND s_Email='" + sEmail.Trim() + "'";
            if (type.Equals("Agency"))
            {
                strSql = "  SELECT n_ContactID FROM dbo.TCstmr_AgencyContact WHERE n_AgencyID=" + nClientID + " AND s_FirstName='" + sName.Trim() + "' AND s_Email='" + sEmail.Trim() + "'";
            }
            return GetIDbySql(strSql);
        }

        #endregion

        #region 客户-财务
        private int AddBill(DataRow _dr)
        {
            int nClientID = GetClientandApplicantIDByName(_dr["客户代码"].ToString(), "Client");
            string strSql = "";
            if (nClientID > 0)
            {
                strSql = "update TCstmr_Client set s_AccountName='" + _dr["单位名称"].ToString() + "',s_TaxCode='" + _dr["纳税人识别号"].ToString() + "',s_RegAddress='" + _dr["注册地址"].ToString() + "',s_RegTel='" + _dr["注册电话"].ToString() + "',s_BankName='" + _dr["开户行名称"].ToString() + "',s_AccountNo='" + _dr["银行账号"].ToString() + "',s_InvoiceTo='" + _dr["发票抬头"].ToString() + "'" +
                    "  where n_ClientID =" + nClientID;
                return InsertbySql(strSql);
            }
            return 0;
        }
        #endregion

        #region 更新客户和申请人关系
        private void updateApplicantandClient()
        {
            string strSql = " UPDATE TCstmr_Applicant SET n_ClientID=(SELECT n_ClientID FROM TCstmr_Client WHERE s_ClientCode=s_AppCode)  WHERE  dbo.TCstmr_Applicant.n_ClientID=-1 AND s_AppCode IS NOT NULL AND s_AppCode!='' AND s_AppCode IN (SELECT s_AppCode FROM dbo.TCstmr_Client WHERE s_ClientCode=s_AppCode) " +
                                 " UPDATE TCstmr_Client SET n_ApplicantID=(SELECT n_AppID FROM TCstmr_Applicant WHERE s_AppCode=TCstmr_Client.s_ClientCode)  WHERE  n_ApplicantID=-1 AND s_ClientCode IS NOT NULL AND s_ClientCode!='' AND s_ClientCode IN (SELECT s_AppCode FROM TCstmr_Applicant WHERE s_AppCode=TCstmr_Client.s_ClientCode) ";
            InsertbySql(strSql);
        }
        #endregion
        #endregion

        #endregion 

        #region 国内
        #region 1.国内-收文数据导入 
        private int InsertFileIn(int i, DataRow dr)
        {
            int result = 0;
            string sNo = dr["我方卷号"].ToString().Trim();
            int numHk = GetIDbyName(sNo, 2);
            if (numHk.Equals(0))
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + numHk + "," + i + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','InsertFileIn','国内-收文数据导入-" + i + "')");
                //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(i + "未找到“我方卷号”为:" + sNo);
                return result;
            }
            if (numHk != 0)
            {
                string remark;
                string sClientGov; //C: 客户 O: 官方

                if (dr["发件人"].ToString().Trim().ToUpper() == "SIPO")
                {
                    sClientGov = "O";
                    string content = "官方来文   " + dr["份"] + "份  ";
                    remark = content;
                }
                else
                {
                    sClientGov = "C";
                    string content = "客户来文   发件人：" + dr["发件人"].ToString().Replace("\r", "").Replace("\n", ""); ;
                    remark = content;
                }
                DateTime insertTime = DateTime.Now.AddMonths(-1);
                string strSql = "";
                try
                {
                    int nClientID = 21;
                    if (sClientGov.Equals("C"))
                    {
                        nClientID =
                            GetIDbySql("  SELECT top 1 n_ClientID FROM TCstmr_Client WHERE s_Name='" +
                                       dr["发件人"].ToString().Trim().Replace("'", "''") + "'");
                        if (nClientID == 0)
                        {
                            strSql =
                                 "INSERT INTO dbo.TCstmr_Client(s_Name) " +
                                 "VALUES('" + dr["发件人"].ToString().Trim().Replace("'", "''").Replace("\r", "").Replace("\n", "") + "-N')";
                            InsertbySql(strSql, i);
                            nClientID =
                            GetIDbySql("  SELECT top 1 n_ClientID FROM TCstmr_Client WHERE s_Name='" +
                                      dr["发件人"].ToString().Trim().Replace("'", "''") + "'");
                        }
                    }
                    string con = dr["内容"].ToString().Replace("'", "''").Trim() == ""
                                     ? "无名称"
                                     : dr["内容"].ToString().Replace("'", "''");

                    strSql =
                       "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,ObjectType,s_Status,dt_EditDate,s_IOType, s_ClientGov,s_SendMethod ,s_Name ,s_Abstact,dt_CreateDate ";
                    string stesql2 = "VALUES  ('国内-收文数据导入" + i + "'," + InFileID + ",'Y','" + insertTime + "','I','" + sClientGov + "','" +
                                     dr["方式"].ToString().Replace("'", "''") + "','" + con + "','" +
                                     remark.Replace("'", "''") + "','" + insertTime + "'";
                    if (dr["发文日"] != null && !string.IsNullOrEmpty(dr["发文日"].ToString()))
                    {
                        strSql += ",dt_SendDate";
                        stesql2 += ",'" + dr["发文日"] + "'";
                    }
                    if (dr["收文日"] != null && !string.IsNullOrEmpty(dr["收文日"].ToString()))
                    {
                        strSql += ",dt_ReceiveDate";
                        stesql2 += ",'" + dr["收文日"] + "'";
                    }
                    if (sClientGov.Equals("C"))
                    {
                        strSql += ",n_ClientID";
                        stesql2 += "," + nClientID + "";
                    }

                    strSql += ")";
                    stesql2 += ")";

                    int insertnum = 0;
                    string strSqlS =
                        "SELECT top 1 n_FileID FROM dbo.T_MainFiles WHERE s_Status='Y' AND s_IOType='I' and  s_ClientGov='" +
                        sClientGov + "' and s_SendMethod='" + dr["方式"].ToString().Replace("'", "''") +
                        "' and s_Name='" + con + "' and s_ClientGov='" + sClientGov + "' and s_Abstact='" +
                        remark.Replace("'", "''") + "' and ObjectType=" + InFileID + " and dt_CreateDate='" + insertTime + "'";
                    if (sClientGov.Equals("C"))
                    {
                        strSqlS += " and n_ClientID=" + nClientID;
                    }
                    if (dr["发文日"] != null && !string.IsNullOrEmpty(dr["发文日"].ToString()))
                    {
                        strSqlS += " and dt_SendDate='" + dr["发文日"] + "'";
                    }
                    if (dr["收文日"] != null && !string.IsNullOrEmpty(dr["收文日"].ToString()))
                    {
                        strSqlS += " and dt_ReceiveDate='" + dr["收文日"] + "'";
                    }
                    strSqlS += " order by n_FileID desc";
                    //int nFileID = GetIDbySql(strSqlS);
                    int nFileID = 0;
                    //if (nFileID > 0)
                    //{
                    //    insertnum = nFileID;
                    //}
                    //else
                    //{
                    //    if (insertnum <= 0)
                    //    {
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = strSql + stesql2;
                        cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                        if (cmd.ExecuteNonQuery() > 0)
                        {
                            nFileID = GetIDbySql(strSqlS);
                            if (nFileID > 0)
                            {
                                insertnum = nFileID;
                            }
                        }
                    }
                    //    } 
                    //}
                    int nGovOfficeID = 0;
                    if (sClientGov.Equals("O")) //s_ClientGov C: 客户 O: 官方
                    {
                        nGovOfficeID = 21; //中国国家知识产权局
                    }
                    if (insertnum > 0)
                    {
                        string sql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                                     "VALUES  (" + numHk + " ," + insertnum + ",'Y')";
                        strSql = "SELECT COUNT(*) AS sumNum FROM dbo.T_FileInCase WHERE n_CaseID=" + numHk +
                                 " AND n_FileID=" + insertnum;
                        int sumNumFileInCase = GetIDbySql(strSql);
                        if (sumNumFileInCase <= 0)
                        {
                            InsertbySql(sql, i);
                        }
                        else
                        {
                            InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + numHk + "," + i + ",'" + sNo + "','" + strSql + "','T_InFiles','国内-收文数据导入-" + i + "')");
                            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(i + "插入T_InFiles表数已存在" + strSql);
                        }

                        strSql = "SELECT COUNT(*) AS sumNum FROM dbo.T_InFiles WHERE n_FileID=" + insertnum +
                                 " and n_GovOfficeID=" + nGovOfficeID + " and s_Distribute='Y'  and s_Note='N'";
                        int sumNumInFiles = GetIDbySql(strSql);
                        if (sumNumInFiles <= 0)
                        {
                            sql =
                                "INSERT INTO dbo.T_InFiles(n_FileID,n_GovOfficeID,dt_TransmitDate,dt_GetCertificatedate,s_Distribute,s_Note)" +
                                "VALUES  (" + nFileID + " ," + nGovOfficeID + ",'" + DateTime.Now + "','" +
                                DateTime.Now + "','Y','N')";
                            InsertbySql(sql, i);
                        }
                    }
                    else
                    {
                        //
                        Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(i + "插入T_MainFiles表数失败" + strSql + stesql2);
                    }
                    result = 1;
                }
                catch (Exception ex)
                {
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(i + "插入T_MainFiles表数据错误，“我方卷号”为:" +
                                                                               dr["我方卷号"].ToString().Trim() +
                                                                               " 错误信息:" + ex.Message);
                }
            }
            return result;
        }

        #endregion

        #region 2.国内-专利数据补充导入 
        private int InsertPatented(int i, DataRow dr)
        {
            string Sql = "";
            int Result = 0;
            string s_No = dr["我方卷号"].ToString().Trim();
            int HkNum = GetIDbyName(s_No, 2);
            if (HkNum.Equals(0))
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + HkNum + "," + i + ",'" + s_No + "','未找到“我方卷号”为:" + s_No + "','InsertPatented','国内-专利数据补充导入-" + i + "')");
                //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("未找到“我方卷号”为:" + s_No);
                return Result;
            }
            if (HkNum != 0)
            {
                #region
                try
                {
                    string No = ResultNo(dr["母案文号"].ToString());
                    string[] ArrayNo = No.Split(',');//原案文号
                    string AppNo = string.Empty;//原案申请号


                    string dtAppDate = string.Empty;//原案申请日
                    //string sDivisionAppNo = string.Empty;//原案申请号s_DivisionAppNo
                    No = ArrayNo[0];
                    AppNo = ArrayNo[1];
                    dtAppDate = ArrayNo[2];

                    //string strSql = "select a.s_AppNo from tcase_base a  left join TPCase_Patent b on  a.n_CaseID=b.n_CaseID left join TPCase_LawInfo c on  b.n_LawID=c.n_ID" +
                    //                 "   where s_caseserial='" + dr["我方卷号"].ToString().Trim() + "'";
                    //object obj = GetTimebySql(strSql);
                    //if (obj != null && !string.IsNullOrEmpty(obj.ToString()))
                    //{
                    //    sDivisionAppNo = obj.ToString();
                    //}
                    int bDivisionalCaseFlag = dr["母案文号"].ToString() == "" ? 0 : 1;
                    Sql = "UPDATE TPCase_Patent set  s_OrigCaseNo='" + No + "', s_OrigAppNo='" + AppNo +
                          "',b_DivisionalCaseFlag=" + bDivisionalCaseFlag;

                    if (bDivisionalCaseFlag != 0)
                    {
                        //if (!string.IsNullOrEmpty(sDivisionAppNo))
                        //{
                        //    Sql += ",s_DivisionAppNo='" + sDivisionAppNo + "'";
                        //}  
                        if (dr["分案申请提交日"] != null && !string.IsNullOrEmpty(dr["分案申请提交日"].ToString()))
                        {
                            Sql += ",dt_DivSubmitDate='" + dr["分案申请提交日"] + "'";
                        }
                    }
                    if (!string.IsNullOrEmpty(dtAppDate))
                    {
                        Sql += ",dt_OrigAppDate='" + dtAppDate + "'";
                    }
                    if (dr["提实审日期"] != null && !string.IsNullOrEmpty(dr["提实审日期"].ToString()))
                    {
                        Sql += ",dt_RequestSubmitDate='" + dr["提实审日期"] + "'";
                    }

                    Sql += " WHERE n_CaseID=" + HkNum;
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = Sql;
                        cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + HkNum + "," + i + ",'" + s_No + "','更新TPCase_Patent信息错误：" + ex.Message + "','TPCase_Patent','国内-专利数据补充导入-" + i + "','" + Sql.Replace("'", "''") + "')");
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("更新TPCase_Patent信息错误：" + ex.Message);
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("更新TPCase_Patent信息错误：" + Sql);
                }
                try
                {
                    Sql = "UPDATE TPCase_LawInfo  set n_ClaimCount='" + dr["权项数"] + "', s_PCTPubNo='" + dr["PCT公开号"] +
                          "'";
                    if (dr["PCT公开日"] != null && !string.IsNullOrEmpty(dr["PCT公开日"].ToString()))
                    {
                        Sql += ",dt_PCTPubDate='" + dr["PCT公开日"] + "'";
                    }
                    if (dr["初审合格日"] != null && !string.IsNullOrEmpty(dr["初审合格日"].ToString()))
                    {
                        Sql += ",dt_FirstCheckDate='" + dr["初审合格日"] + "'";
                    }
                    if (dr["进入实审发文日"] != null && !string.IsNullOrEmpty(dr["进入实审发文日"].ToString()))
                    {
                        Sql += ",dt_SubstantiveExamDate='" + dr["进入实审发文日"] + "'";
                    }
                    Sql += "  WHERE n_ID IN (SELECT n_LawID FROM TPCase_Patent  WHERE n_CaseID=" + HkNum + ")";
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = Sql;
                        cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + HkNum + "," + i + ",'" + s_No + "','更新TPCase_LawInfo信息错误：" + ex.Message + "','TPCase_LawInfo','国内-专利数据补充导入-" + i + "','" + Sql.Replace("'", "''") + "')");
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("更新TPCase_LawInfo信息错误：" + ex.Message +
                                                                               "  " + Sql);
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("更新TPCase_LawInfo信息错误：" + Sql);
                }

                #endregion

                #region  增加要求

                if (dr["年费备注"] != null && !string.IsNullOrEmpty(dr["年费备注"].ToString()))
                {
                    InsertSemand(dr["年费备注"].ToString().Replace("'", "''").Replace("\r", "").Replace("\n", ""), HkNum, i);
                }
                if (dr["客户指示提实审"] != null && !string.IsNullOrEmpty(dr["客户指示提实审"].ToString()) && dr["客户指示提实审"].ToString().Trim().ToUpper().Equals("Y"))
                {
                    //添加GA01 要求
                    string sql = "select n_ID from TCode_Demand where s_sysdemand='GA01'";
                    DataTable table = GetDataTablebySql(sql);
                    if (table.Rows.Count > 0)
                    {
                        sql = "select n_ID from T_Demand where n_CodeDemandID=" + table.Rows[0]["n_ID"].ToString();
                        int num = GetIDbySql(sql);

                        InsertCaseE(num, HkNum.ToString());
                    }
                }
                if (dr["客户指示"] != null && !string.IsNullOrEmpty(dr["客户指示"].ToString()))
                {
                    //客户指示增加到个案要求内
                    InsertCaseD(dr["客户指示"].ToString().Replace("'", "''").Replace("\r", "").Replace("\n", ""), HkNum);
                }
                #endregion

                #region 增加备注  权项数  字数 增加到自定义属性

                //if (dr["权项数"] != null && !string.IsNullOrEmpty(dr["权项数"].ToString()))
                //{
                //    //读取劝降书自定义属性ID
                //    string strSql = " SELECT n_ID FROM TCode_CaseCustomField WHERE  s_IPType='P' AND s_IsActive='Y' AND s_CustomFieldName IN ('权项数')";
                //    int nID = GetIDbySql(strSql);
                //    if (nID > 0)
                //    {
                //        InorUpdateCaseCustomField(HkNum, nID, dr["权项数"].ToString());
                //    }
                //}
                //if (dr["字数"] != null && !string.IsNullOrEmpty(dr["字数"].ToString()))
                //{
                //    string strSql = " SELECT n_ID FROM TCode_CaseCustomField WHERE  s_IPType='P' AND s_IsActive='Y' AND s_CustomFieldName IN ('字数（说明书+附图）')";
                //    int nID = GetIDbySql(strSql);
                //    if (nID > 0)
                //    {
                //        InorUpdateCaseCustomField(HkNum, nID, dr["字数"].ToString());
                //    }
                //}
                #endregion

                #region 用户与案件关联

                string Applicant = dr["申请人"].ToString();

                if (dr["第二委托人"] != null && !string.IsNullOrEmpty(dr["第二委托人"].ToString()) &&
                    dr["第二委托人"].ToString().Trim() != "0")
                {
                    if (!Applicant.Equals(dr["第二委托人"].ToString()))
                    {
                        InsertTCaseClients(dr["第二委托人"].ToString(), HkNum, i, "国内-专利数据补充导入");
                    }
                }
                if (dr["第三委托人"] != null && !string.IsNullOrEmpty(dr["第三委托人"].ToString()) &&
                    dr["第三委托人"].ToString().Trim() != "0")
                {
                    if (!Applicant.Equals(dr["第三委托人"].ToString()))
                    {
                        InsertTCaseClients(dr["第三委托人"].ToString(), HkNum, i, "国内-专利数据补充导入");
                    }
                }
                if (dr["第四委托人"] != null && !string.IsNullOrEmpty(dr["第四委托人"].ToString()) &&
                    dr["第四委托人"].ToString().Trim() != "0")
                {
                    if (!Applicant.Equals(dr["第四委托人"].ToString()))
                    {
                        InsertTCaseClients(dr["第四委托人"].ToString(), HkNum, i, "国内-专利数据补充导入");
                    }
                }
                if (dr["第五委托人"] != null && !string.IsNullOrEmpty(dr["第五委托人"].ToString()) &&
                    dr["第五委托人"].ToString().Trim() != "0")
                {
                    if (!Applicant.Equals(dr["第五委托人"].ToString()))
                    {
                        InsertTCaseClients(dr["第五委托人"].ToString(), HkNum, i, "国内-专利数据补充导入");
                    }
                }

                #endregion

                #region 增加案件处理人信息

                if (dr["翻译人"] != null && !string.IsNullOrEmpty(dr["翻译人"].ToString()) && dr["翻译人"].ToString().Trim() != "0")
                {
                    InsertUser(dr["翻译人"].ToString(), HkNum, i, "代理部-新申请阶段-翻译人", "国内-专利数据补充导入");
                }
                if (dr["一校"] != null && !string.IsNullOrEmpty(dr["一校"].ToString()) && dr["一校"].ToString().Trim() != "0")
                {
                    InsertUser(dr["一校"].ToString(), HkNum, i, "代理部-新申请阶段--一校", "国内-专利数据补充导入");
                    InsertUser(dr["一校"].ToString(), HkNum, i, "代理部-新申请阶段-办案人", "国内-专利数据补充导入");
                }
                if (dr["二校"] != null && !string.IsNullOrEmpty(dr["二校"].ToString()) && dr["二校"].ToString().Trim() != "0")
                {
                    InsertUser(dr["二校"].ToString(), HkNum, i, "代理部-新申请阶段-二校", "国内-专利数据补充导入");
                }
                string ForeignAgentCode = dr["对外代理人代码"].ToString();
                string ActualAgentCode = dr["实际代理人代码"].ToString();
                if (!string.IsNullOrEmpty(ActualAgentCode) && !string.IsNullOrEmpty(ForeignAgentCode))
                {
                    if (dr["实际代理人代码"] != null && !string.IsNullOrEmpty(dr["实际代理人代码"].ToString()) &&
                        dr["实际代理人代码"].ToString().Trim() != "0")
                    {
                        InsertUser(dr["实际代理人代码"].ToString(), HkNum, i, "代理部-新申请阶段-办案人", "国内-专利数据补充导入");
                    }
                }
                else
                {
                    if (dr["实际代理人代码"] != null && !string.IsNullOrEmpty(dr["实际代理人代码"].ToString()) && dr["实际代理人代码"].ToString().Trim() != "0")
                    {
                        InsertUser(dr["实际代理人代码"].ToString(), HkNum, i, "代理部-新申请阶段-办案人", "国内-专利数据补充导入");
                    }
                    if (dr["对外代理人代码"] != null && !string.IsNullOrEmpty(dr["对外代理人代码"].ToString()) &&
                        dr["对外代理人代码"].ToString().Trim() != "0")
                    {
                        InsertUser(dr["对外代理人代码"].ToString(), HkNum, i, "对外代理人", "国内-专利数据补充导入");
                    }
                }

                #endregion

                Result = 1;
            }
            return Result;
        }

        private void InsertSemand(string s_Title, int n_CaseID, int i)
        {
            string Sql = "SELECT n_ID  FROM T_Demand where  s_Title='" + s_Title.ToString().Replace("'", "''") +
                         "' and n_CaseID=" + n_CaseID;
            int n_ID = GetIDbySql(Sql);
            if (n_ID <= 0)
            {
                Sql = "INSERT INTO dbo.T_Demand(s_sourcetype1,dt_EditDate,s_Title,dt_CreateDate,n_CaseID)" +
                      "VALUES  ('年费备注','" + DateTime.Now + "','" + s_Title.ToString().Replace("'", "''") + "','" + DateTime.Now +
                      "'," + n_CaseID + ")";
                int ResultNum = InsertbySql(Sql, i);
                if (ResultNum.Equals(0))
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + n_CaseID + "," + i + ",'','增加年费备注失败[InsertSemand]','T_Demand','国内-专利数据补充导入-" + i + "','" + Sql.Replace("'", "''") + "')");
                }
            }
            else
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("已存在年费备注[InsertSemand]:" + n_CaseID + ":" + s_Title + "==" + Sql.Replace("'", "''"));
            }
        }
        //自定义属性
        //private void InorUpdateCaseCustomField(int HkNum, int nID, string sValue)
        //{
        //    string strSql = "SELECT n_ID FROM TCase_CaseCustomField  WHERE n_CaseID=" + HkNum + " AND n_FieldCodeID=" + nID;
        //    int nCaseFieldID = GetIDbySql(strSql);
        //    if (nCaseFieldID > 0)
        //    {
        //        strSql = "update TCase_CaseCustomField set s_Value='" + sValue + "' WHERE n_CaseID=" + HkNum + " AND n_FieldCodeID=" + nID;
        //    }
        //    else
        //    {
        //        strSql = "INSERT INTO dbo.TCase_CaseCustomField( n_CaseID, n_FieldCodeID, s_Value )VALUES (" + HkNum + "," + nID + ",'" + sValue + "')";
        //    }
        //    if (InsertbySql(strSql, 0) <= 0)
        //    {
        //        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + HkNum + ",0,'InorUpdateCaseCustomField','自定义属性数据添加失败[InorUpdateCaseCustomField]','TCase_CaseCustomField','国内-专利数据补充导入-','" + strSql.Replace("'", "''") + "')");
        //        Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("自定义属性数据添加失败[InorUpdateCaseCustomField]:" + strSql);

        //    }
        //}

        private void InsertCaseE(int nID, string nCaseID)
        {
            //根据系统要求代码查询主题和描述
            string strl = "select n_ID,s_IPtype,s_Title,s_Description,s_Creator,n_DemandType,n_SysDemandID,n_CodeDemandID,s_ModuleType,s_sysDemand from T_Demand WHERE n_ID=" + nID;

            DataTable newTable = GetDataTablebySql(strl);

            string s_IPType = string.Empty;
            string title = string.Empty;
            string description = string.Empty;
            string sCreator = string.Empty;
            string n_DemandType = string.Empty;
            string sModuleType = string.Empty;
            string s_sysDemand = string.Empty;
            string n_SysDemandID = string.Empty;
            //string 
            if (newTable.Rows.Count > 0)
            {
                for (int k = 0; k < newTable.Rows.Count; k++)
                {
                    n_DemandType = newTable.Rows[k]["n_DemandType"].ToString();
                    s_IPType = newTable.Rows[k]["s_IPtype"].ToString();
                    title = newTable.Rows[k]["s_Title"].ToString();
                    description = newTable.Rows[k]["s_Description"].ToString();
                    sCreator = newTable.Rows[k]["s_Creator"].ToString();
                    sModuleType = newTable.Rows[k]["s_ModuleType"].ToString();
                    s_sysDemand = newTable.Rows[k]["s_sysDemand"].ToString();
                    n_SysDemandID = newTable.Rows[k]["n_SysDemandID"].ToString();
                    //if (sModuleType == "ClientApplicant")
                    //{
                    //    sModuleType = "Applicant";
                    //}
                    string Sql = "INSERT INTO dbo.T_Demand(s_sourcetype1,s_ModuleType,s_Title,s_Description,s_Creator,s_Editor,s_IPType,s_SysDemand,n_DemandType,dt_EditDate,dt_CreateDate,n_SysDemandID,n_CodeDemandID,s_SourceModuleType,n_CaseID)" +
                                 "VALUES  ('国内-专利数据补充导入--客户指示提实审','Case','" + title + "','" + description + "','" + sCreator + "','" + sCreator + "','" + s_IPType + "','" + s_sysDemand + "','" + n_DemandType + "','" + DateTime.Now + "','" + DateTime.Now + "'," + n_SysDemandID + "," + n_SysDemandID + ",'Case','" + nCaseID + "')";

                    //查询是否存在此案件要求要求
                    string strSql = "select n_ID from T_Demand where s_ModuleType='Case'  and n_CaseID=" + nCaseID + " and  n_codeDemandID=" + n_SysDemandID;
                    DataTable Table = GetDataTablebySql(strSql);

                    if (Table.Rows.Count <= 0)
                    {
                        InsertbySql(Sql, 0);
                    }
                    else
                    {
                        string updareSql = "update T_Demand set s_SourceModuleType='" + sModuleType + "' where n_ID=" + int.Parse(Table.Rows[0]["n_ID"].ToString());
                        InsertbySql(updareSql, 0);
                    }
                }
            }
            else
            {
                string sql = "select * from TCode_Demand where s_sysdemand='GA01'";
                newTable = GetDataTablebySql(sql);
                for (int k = 0; k < newTable.Rows.Count; k++)
                {

                    s_IPType = newTable.Rows[k]["s_IPtype"].ToString();
                    title = newTable.Rows[k]["s_Title"].ToString();
                    description = newTable.Rows[k]["s_Description"].ToString();
                    sCreator = newTable.Rows[k]["s_Creator"].ToString();
                    s_sysDemand = newTable.Rows[k]["s_sysDemand"].ToString();

                    string Sql = "INSERT INTO dbo.T_Demand(s_sourcetype1,s_ModuleType,s_Title,s_Description,s_Creator,s_Editor,s_IPType,s_SysDemand,n_DemandType,dt_EditDate,dt_CreateDate,s_SourceModuleType,n_CaseID)" +
                                 "VALUES  ('国内-专利数据补充导入--客户指示提实审','Case','" + title + "','" + description + "','" + sCreator + "','" + sCreator + "','" + s_IPType + "','" + s_sysDemand + "','" + n_DemandType + "','" + DateTime.Now + "','" + DateTime.Now + "','Case','" + nCaseID + "')";

                    //查询是否存在此案件要求要求
                    string strSql = "select n_ID from T_Demand where s_ModuleType='Case'  and n_CaseID=" + nCaseID + " and  s_sourcetype1='国内-专利数据补充导入--客户指示提实审' and s_Title='" + title + "'";
                    DataTable Table = GetDataTablebySql(strSql);

                    if (Table.Rows.Count <= 0)
                    {
                        InsertbySql(Sql, 0);
                    }
                    else
                    {
                        string updareSql = "update T_Demand set s_SourceModuleType='" + sModuleType + "' where n_ID=" + int.Parse(Table.Rows[0]["n_ID"].ToString());
                        InsertbySql(updareSql, 0);
                    }
                }
            }
        }

        private void InsertCaseD(string title, int nCaseID)//
        {
            string selctSql = " select n_ID from TFCode_DemandType WHERE s_Name='其他'";
            int n_DemandType = GetIDbySql(selctSql);//n_DemandType

            string Sql = "INSERT INTO dbo.T_Demand(s_sourcetype1,s_ModuleType,s_Title,s_Description,s_Creator,s_Editor,s_IPType,s_SysDemand,n_DemandType,dt_EditDate,dt_CreateDate,s_SourceModuleType,n_CaseID)" +
                                    "VALUES  ('国内-专利数据补充导入---客户指示','Case','" + title + "','','administrator','administrator','P','','" + n_DemandType + "','" + DateTime.Now + "','" + DateTime.Now + "','Case','" + nCaseID + "')";

            //查询是否存在此案件要求要求
            string strSql = "select n_ID from T_Demand where s_ModuleType='Case'  and n_CaseID=" + nCaseID + " and  s_Title='" + title.Replace("'", "''") + "'";
            DataTable Table = GetDataTablebySql(strSql);

            if (Table.Rows.Count <= 0)
            {
                InsertbySql(Sql, 0);
            }
        }

        private string ResultNo(string sOrigCaseNo)
        {
            string strSql = "select s_OrigCaseNo from tcase_base a  left join TPCase_Patent b on  a.n_CaseID=b.n_CaseID left join TPCase_LawInfo c on  b.n_LawID=c.n_ID" +
                "   where s_caseserial='" + sOrigCaseNo + "'";
            object obj = GetTimebySql(strSql);
            if (obj != null && !string.IsNullOrEmpty(obj.ToString()))
            {
                return ResultNo(obj.ToString());
            }
            else
            {
                strSql = "select a.s_AppNo,a.dt_AppDate from tcase_base a  left join TPCase_Patent b on  a.n_CaseID=b.n_CaseID left join TPCase_LawInfo c on  b.n_LawID=c.n_ID" +
                "   where s_caseserial='" + sOrigCaseNo + "'";
                DataTable table = GetDataTablebySql(strSql);
                if (table.Rows.Count > 0)
                {
                    sOrigCaseNo += "," + table.Rows[0]["s_AppNo"].ToString() + "," + table.Rows[0]["dt_AppDate"].ToString();
                }
                else
                {
                    sOrigCaseNo += ",,";
                }
                return sOrigCaseNo;
            }
        }
        #endregion

        #region 3.国内-OA数据补充导入 
        private int OA(int i, DataRow dr)
        {
            int result = 0;
            string s_No = dr["我方卷号"].ToString().Trim();
            int hkNum = GetIDbyName(s_No, 2);
            if (hkNum.Equals(0))
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + i + ",'" + s_No + "','未找到“我方卷号”为:" + s_No + "','OA','国内-OA数据补充导入-" + i + "')");
                //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(i + "未找到“我方卷号”为:" +
                //                                                           dr["我方卷号"].ToString().Trim());
                return result;
            }
            if (hkNum != 0)
            {
                string nameType;
                string sClientGov;
                if (!string.IsNullOrEmpty(dr["转发客户日"].ToString().Trim()))
                {
                    nameType = dr["类型"] == null ? "OA转发客户" : dr["类型"] + "转客户";
                    sClientGov = "C";
                    InsertIntos(i, dr, sClientGov, nameType, hkNum, dr["转发客户日"].ToString().Trim());

                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(0," + i + ",0,'转发客户日-时间格式为空，无需导入','','国内-OA数据补充导入-" + i + "','')");
                }
                if (!string.IsNullOrEmpty(dr["账单及代理报告发出日"].ToString().Trim()))
                {
                    nameType = "报告及账单发客户";
                    sClientGov = "C";
                    InsertIntos(i, dr, sClientGov, nameType, hkNum, dr["账单及代理报告发出日"].ToString().Trim());
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(0," + i + ",0,'账单及代理报告发出日-时间格式为空，无需导入','','国内-OA数据补充导入-" + i + "','')");
                }
                if (!string.IsNullOrEmpty(dr["答复日"].ToString().Trim()))
                {
                    nameType = "发官方文";
                    sClientGov = "O";
                    InsertIntos(i, dr, sClientGov, nameType, hkNum, dr["答复日"].ToString().Trim());
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(0," + i + ",0,'答复日-时间格式为空，无需导入','','国内-OA数据补充导入-" + i + "','')");
                }
                if (!string.IsNullOrEmpty(dr["OA收到日"].ToString().Trim()))
                {
                    nameType = "官方来文";
                    sClientGov = "O";
                    InsertIntos(i, dr, sClientGov, nameType, hkNum, dr["OA收到日"].ToString().Trim());
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(0," + i + ",0,'OA收到日-时间格式为空，无需导入','','国内-OA数据补充导入-" + i + "','')");
                }
                result = 1;
            }
            return result;
        }

        private void InsertIntos(int i, DataRow dr, string sClientGov, string nameType, int numHk, string time)
        {
            string remark = dr["OA延期"] == null ? "" : dr["OA延期"].ToString();
            string people = dr["代理人"] == null ? "" : dr["代理人"].ToString();

            remark = remark == "" ? "" : "OA延期:" + remark;
            people = people == "" ? "" : "代理人:" + people;
            DateTime insertTime = DateTime.Now.AddMonths(-1);

            string strSql =
                "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,ObjectType,dt_EditDate,s_Name,s_Abstact,dt_CreateDate,s_ClientGov,s_IOType,dt_SendDate,s_Status)" +
                "VALUES  ('国内-OA数据补充导入" + i + "'," + OutFileID + ",'" + insertTime + "','" + nameType.Replace("'", "''") + "','" +
                remark.Replace("'", "''") + "  " + people.Replace("'", "''") + "','" + insertTime + "','" + sClientGov + "','O','" + time + "','Y')";

            string strSqlS =
                "SELECT top 1 n_FileID FROM dbo.T_MainFiles WHERE  s_IOType='O' AND s_Status='Y' and  s_ClientGov='" +
                sClientGov +
                "' and s_Name='" + nameType.Replace("'", "''") + "' and s_Abstact='" +
                remark.Replace("'", "''") + "  " + people.Replace("'", "''") + "' and dt_SendDate='" + time +
                "' and ObjectType=" + OutFileID + " and dt_CreateDate='" + insertTime + "' order by n_FileID desc ";
            int insertnum = 0;
            int nFileID2 = 0;// GetIDbySql(strSqlS);
            //if (nFileID2 > 0)
            //{
            //    insertnum = nFileID2;
            //}
            //else
            //{
            if (InsertbySql(strSql, i) > 0)
            {
                nFileID2 = GetIDbySql(strSqlS);
                if (nFileID2 > 0)
                {
                    insertnum = nFileID2;
                }
            }
            //}
            if (insertnum > 0)
            {
                int fCount = 0;
                string sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_FileInCase WHERE n_FileID=" + insertnum +
                             " and n_CaseID=" + numHk;
                int sumFileInCase = GetIDbySql(sql);
                if (sumFileInCase <= 0)
                {
                    sql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                          "VALUES  (" + numHk + " ," + insertnum + ",'Y')";
                    int fileInCase = InsertbySql(sql, i); //记录文件、案件与程序的关系表 
                    if (fileInCase <= 0)
                    {
                        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + numHk + "," + i + ",0,'T_FileInCase插入数据错误','T_FileInCase','国内-OA数据补充导入-" + i + "','" + sql.Replace("'", "''") + "')");
                        //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(i + "T_FileInCase插入数据错误:" + sql);
                    }
                    fCount = fileInCase;
                }
                else
                {
                    fCount = sumFileInCase;
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + numHk + "," + i + ",0,'已存在案件和文件关系，无需插入','T_FileInCase','国内-OA数据补充导入-" + i + "','" + sql.Replace("'", "''") + "')");
                }
                int nGovOfficeID = 0;
                if (sClientGov.Equals("O")) //s_ClientGov C: 客户 O: 官方
                {
                    nGovOfficeID = 21; //中国国家知识产权局
                }
                if (fCount > 0)
                {
                    if (!nameType.Equals("官方来文"))
                    {
                        string sqlS =
                            "INSERT INTO dbo.T_OutFiles( n_FileID ,n_CheckedOutBy , n_GovOfficeID , s_FileStatus, dt_StatusDate ,dt_WriteDate ," +
                            "n_WriterID , n_SubmiterID ,  n_PrintNum , n_PageNum ,n_ReFileID  ,n_Count ,s_FileType ,n_LatestCheckInfoID)" +
                            "VALUES  (" + insertnum + ",0 ," + nGovOfficeID + " ,'W' ,'" + DateTime.Now + "' ,'" +
                            DateTime.Now + "',0 ,0 ,1,0 ,0,0 ,'new',0 )";
                        sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_OutFiles WHERE n_FileID=" + insertnum;
                        int sumcount = GetIDbySql(sql);
                        if (sumcount <= 0)
                        {
                            int outFiles = InsertbySql(sqlS, i);
                            if (outFiles <= 0)
                            {
                                //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(i + "T_OutFiles插入数据错误:" +
                                //                                                           sqlS);
                                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + numHk + "," + i + ",0,'T_OutFiles插入数据错误','T_OutFiles','国内-OA数据补充导入-" + i + "','" + sqlS.Replace("'", "''") + "')");
                            }
                        }
                    }
                    else
                    {
                        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + numHk + "," + i + ",0,'官方来文无需插入','T_OutFiles','国内-OA数据补充导入-" + i + "','')");
                    }
                }
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(0," + i + ",0,'未查询到相关的文件往来','T_MainFiles','国内-OA数据补充导入-" + i + "','1." + strSqlS.Replace("'", "''") + "  2." + strSql.Replace("'", "''") + "')");
            }
        }

        #endregion

        #region 3-1.案件处理人  国内-OA数据补充导入(辅表)-已翻译代理人

        private int InsertCaseAttorney(DataTable table, int i, DataRow dr)
        {
            int result = 0;
            string sNo = dr["我方卷号"].ToString().Trim();
            int hkNum = GetIDbyName(sNo, 2);
            if (hkNum.Equals(0))
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + i + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','InsertCaseAttorney','国内-OA数据补充导入(辅表)-已翻译代理人-" + i + "')");
                //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(i + "未找到“我方卷号”为:" + sNo);
            }
            else
            {
                result = 1;
                if (!string.IsNullOrEmpty(dr["代理人"].ToString().Trim()))
                {
                    //如果多个处理人进行循环处理
                    string[] ArryUser = dr["代理人"].ToString().Split(' ');
                    if (ArryUser.Length > 0)
                    {
                        if (!string.IsNullOrEmpty(ArryUser[0]))
                        {
                            InsertUser(ArryUser[0].Trim().Replace("　", ""), hkNum, i, "代理部-OA阶段-办案人", "国内-OA数据补充导入(辅表)-已翻译代理人");
                        }
                    }
                }
            }
            return result;
        }

        #endregion

        #region 4.国内-年费
        private int InsertFee(int i, DataRow dr)
        {
            int result = 0;
            int year = 0;
            string sNo = dr["案件文号"].ToString().Trim();
            int hkNum = GetIDbyName(sNo, 2);
            if (hkNum.Equals(0))
            {
                //未找到“我方卷号” 
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + i + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','InsertFee','国内-年费-" + i + "')");
                return result;
            }
            else
            {
                result = 1;
                //专利案件申请日  
                string Sql = "SELECT dt_AppDate FROM TCase_Base WHERE n_CaseID=" + hkNum;
                object time = GetTimebySql(Sql);
                if (time != null && time.ToString() != "")
                {
                    year = DateTime.Parse(time.ToString()).Year;
                }

                //专利案件国家  
                //Sql = "SELECT n_RegCountry FROM TCase_Base WHERE n_CaseID=" + hkNum;
                //GetIDbySql(Sql);

                //专利类型
                Sql = "SELECT n_PatentTypeID FROM TPCase_Patent WHERE n_CaseID=" + hkNum;
                int n_PatentTypeID = GetIDbySql(Sql);

                //读取年费标准 
                Sql = "SELECT  n_YearNo,n_OfficialFee FROM TCode_AnnualFee WHERE n_PatentType=" + n_PatentTypeID +
                      " ORDER BY n_YearNo asc";
                DataTable tableYearNo = GetDataTablebySql(Sql);

                //读取年费标准年数
                Sql = " SELECT  count(*) as YearSum  FROM TCode_AnnualFee WHERE n_PatentType=" + n_PatentTypeID;
                int YearSum = GetIDbySql(Sql);
                if (dr["下次年费年度"].ToString() != "")
                {
                    DateTime Next = DateTime.Parse(dr["下次年费年度"].ToString());
                    int NextTime = Next.Year;

                    int Sumnum = year + YearSum - NextTime;
                    int Start = YearSum - Sumnum + 1;
                    if (tableYearNo != null)
                    {
                        for (int iS = Start; iS < YearSum + 1; iS++)
                        {
                            //查询是否存在当前年份的年费
                            Sql = "SELECT n_AnnualFeeID FROM T_AnnualFee WHERE n_CaseID=" + hkNum + " AND n_YearNo=" + iS;
                            int n_AnnualFeeID = GetIDbySql(Sql);
                            if (n_AnnualFeeID > 0)
                            {
                                //如果存在当前年的年费不做处理
                                Sql = "update T_AnnualFee set dt_OfficialShldPayDate='" + Next + "',dt_AlarmDate='" +
                                      Next.AddMonths(-2) + "' WHERE n_AnnualFeeID=" + n_AnnualFeeID;
                                int numS = InsertbySql(Sql, i);
                                if (numS == 0)
                                {
                                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + i + ",'" + sNo + "','修改年费数据插入错误：" + Sql.Replace("'", "''") + "','T_AnnualFee','国内-年费-" + i + "')");
                                    //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(i + "年费数据插入错误SQL:" + Sql);
                                }
                            }
                            else
                            {
                                //循环产生年费记录
                                Sql =
                                    "INSERT INTO dbo.T_AnnualFee( n_CaseID ,n_YearNo , s_Status , s_PayMode , s_StatusOrder , n_ChargeCurrency , n_ChargeOFee , n_OfficialCurrency , n_OfficialFee ,  s_IsOfficialDisc ,s_OfficialDiscStyle , dt_OfficialShldPayDate ,dt_AlarmDate,s_IsActive ,dt_CreateDate ,dt_EditDate)" +
                                    "VALUES  (" + hkNum + "," + (iS) + ",'XXNNN','AX','123' ,8 ,'" +
                                    tableYearNo.Rows[iS - 1]["n_OfficialFee"] + "' ,8 ,'" +
                                    tableYearNo.Rows[iS - 1]["n_OfficialFee"] + "','N' ,'2' , '" + Next + "','" +
                                    Next.AddMonths(-2) + "','Y','" + DateTime.Now + "','" + DateTime.Now + "')";

                                int numS = InsertbySql(Sql, i);
                                result = numS;
                                if (numS == 0)
                                {
                                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + i + ",'" + sNo + "','增加年费数据插入错误：" + Sql.Replace("'", "''") + "','T_AnnualFee','国内-年费-" + i + "')");
                                    //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(i + "年费数据插入错误SQL:" + Sql);
                                }
                            }
                            Next = Next.AddYears(1);
                        }
                    }
                }
                else
                {
                    //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(i + "下次年费年度为空，无法导入" + sNo);
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + i + ",'" + sNo + "','下次年费年度为空,无法导入','T_AnnualFee','国内-年费-" + i + "')");
                }
            }
            return result;
        }

        #endregion

        #region 5.国内-优先权
        private int TPCasePriority(int rowid, DataRow dr)
        {
            int result = 0;
            #region
            string sNo = dr["我方卷号"].ToString().Trim();
            int Country = GetIDbyName(dr["优先权国家"].ToString().Trim(), 1);
            int HKNum = GetIDbyName(sNo, 2);
            if (HKNum.Equals(0))
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + HKNum + "," + rowid + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','TPCasePriority','国内-优先权-" + rowid + "')");
                return result;
            }
            else
            {
                result = 1;
                InsertTPCasePriority(HKNum, Country, dr, "国内-优先权", rowid);
            }
            UpdateSeq(HKNum);

            if (dr["优先权国家"].ToString().Trim().Equals("中国")) //B为主案
            {
                string strSql =
                    "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='国内优先权' AND s_MasterName='国内案' AND s_SlaveName='国外案' AND s_IPType='P'";
                int n_ID = GetIDbySql(strSql);
                InsertInto(dr["优先权号"].ToString().Trim(), HKNum, n_ID, rowid);
            }
            #endregion
            return result;
        }

        private void InsertInto(string No, int HKNum, int n_ID, int rowid)
        {
            int caseID = GetIDbyName(No, 7);
            if (caseID > 0)
            {
                string strSql = "SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where n_CaseIDA=" + HKNum + " and n_CaseIDB=" + caseID + " and n_CodeRelativeID=" + n_ID;
                int NUMS =
                    GetIDbySql(strSql);
                if (NUMS <= 0 && HKNum > 0)
                {
                    //InsertTCaseCaseRelative(HKNum, caseID, n_ID, 0);
                    strSql = " INSERT INTO  dbo.TCase_CaseRelative ( n_CaseIDA ,  n_CaseIDB , dt_CreateDate , dt_EditDate , s_MasterSlaveRelation , n_CodeRelativeID )" +
                          " VALUES  ( " + HKNum + " , " + caseID + " ,  GETDATE() , GETDATE() ,  0,  " + n_ID + ")";
                    //int Num = InsertbySql(strSql, 0);
                    if (InsertbySql(strSql, 0) <= 0)
                    {
                        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + HKNum + "," + rowid + ",'" + caseID + "','案件关系插入数据失败','TCase_CaseRelative','国内-优先权-" + rowid + "','" + strSql.Replace("'", "''") + "')");
                    }
                }
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + HKNum + "," + rowid + ",'" + caseID + "','优先权号未查到，无法添加，优先权号:" + No + "','TCase_CaseRelative','国内-优先权-" + rowid + "','')");
                // Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(i + "优先权号未查到，无法添加");
            }
        }

        #endregion

        #endregion

        #region 国外
        #region 6.国外-法律信息及日志表
        private int TCaseLaw(int rowid, DataRow dr)
        {
            int result = 0;
            string sNo = dr["我方卷号"].ToString().Trim();
            int hkNum = GetIDbyName(sNo, 2);
            if (hkNum.Equals(0))
            {
                //未找到“我方卷号” 
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','TCaseLaw','国内-法律信息及日志表-" + rowid + "')");
                return result;
            }
            DateTime insertTime = DateTime.Now.AddMonths(-1);
            if (dr["项目"] != null && dr["项目"].ToString().Trim() != "办登日" && dr["项目"].ToString().Trim() != "IDS号" &&
                dr["项目"].ToString().Trim() != "IDS日" && dr["项目"].ToString().Trim() != "驳回日" &&
                dr["项目"].ToString().Trim() != "复审日")
            {
                #region
                string sIoType = string.Empty; //文件类型
                string sClientGov = string.Empty; //客户 官方
                string type = string.Empty;
                int nAgencyID = 0;
                #region 判断来文类型

                if (dr["项目"] != null && dr["项目"].ToString().Trim() == "官方发文")//官方发文   		官方来文
                {
                    type = "来文";
                    sIoType = "I"; 
                    sClientGov = "C";
                }

                if (dr["项目"] != null &&
                   ( dr["项目"].ToString().Trim() == "申请人来文" || dr["项目"].ToString().Trim() == "委托人来文"))// 申请人来文、委托人来文		    客户来文
                {
                    type = "来文";
                    sIoType = "I";
                    sClientGov = "O";
                }
                if (dr["项目"] != null &&
                    (dr["项目"].ToString().Trim() == "我方递交官方文件" || dr["项目"].ToString().Trim() == "我方给委托人文" || dr["项目"].ToString().Trim() == "我方给申请人来文"))//我方给申请人来文、我方给委托人文	发客户文
                { //我方递交官方文件	发官方文
                    type = "发文";
                    sIoType = "O";
                    sClientGov = dr["项目"].ToString().Trim() != "我方递交官方文件" ? "C" : "O";
                }
                string strSql;
                if (dr["项目"] != null && dr["项目"].ToString().Trim() == "我方给外代理文")//我方给外代理文   		发代理机构文
                {
                    type = "发文";
                    sIoType = "O";
                    sClientGov = "C";
                    strSql = " SELECT n_CoopAgencyToID  FROM dbo.TCase_Base WHERE n_CaseID=" + hkNum;
                    nAgencyID =GetIDbySql(strSql);
                    if(nAgencyID<=0)
                    {
                        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','项目类型：我方给外代理文，未查到案件代理机构','TCaseLaw','国内-法律信息及日志表-" + rowid + "')");
                    }
                }
                string sClientType = "";
                string sNote = "";
                if (dr["项目"] != null && dr["项目"].ToString().Trim() == "外代理来文")//外代理来文   		外代理来文（系统目前不支持） 	
                {
                    type = "来文";
                    sIoType = "I";
                    sClientGov = "O";
                    sClientType = "O";
                    sNote = "[来文类型：外代理]";
                }
                #endregion

                #region
                strSql = "";
                if (type == "来文" || type == "发文")
                {

                    if (type == "来文")
                    {
                        strSql =
                            "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,s_Abstact,ObjectType,s_SendMethod,dt_EditDate,s_Name,dt_ReceiveDate,dt_CreateDate,s_IOType,s_ClientGov,s_Status,s_ClientType)" +
                            "VALUES  ('国外-法律信息及日志表"+rowid+"',''," + InFileID + ",'其他','" + insertTime + "','" +
                            dr["项目"].ToString().Replace("'", "''") + "  " +
                            dr["内容"].ToString().Replace("'", "''") + "','" + dr["记录日"] + "','" + insertTime + "','" +
                            sIoType + "','" + sClientGov + "','Y','"+sClientType+"')";
                    }
                    else if (type == "发文")
                    {
                        strSql =
                            "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,s_Abstact,ObjectType,s_SendMethod,dt_EditDate,s_Name,dt_SendDate,dt_CreateDate,s_IOType,s_ClientGov,s_Status)" +
                            "VALUES  ('国外-法律信息及日志表"+rowid+"',''," + OutFileID + ",'其他','" + insertTime + "','" +
                            dr["项目"].ToString().Replace("'", "''") + "  " +
                            dr["内容"].ToString().Replace("'", "''") + "','" + dr["记录日"] + "','" + insertTime + "','" +
                            sIoType + "','" + sClientGov + "','Y')   ";
                    }
                    if (!string.IsNullOrEmpty(strSql))
                    {
                        InsertbySql(strSql, 0);
                    }
                    int insertnum = 0;
                    string strSqlS = "SELECT top 1  n_FileID FROM dbo.T_MainFiles WHERE s_Status='Y' AND  s_Name='" +
                                     dr["项目"].ToString().Replace("'", "''") + "  " +
                                     dr["内容"].ToString().Replace("'", "''") + "' and s_ClientGov='" + sClientGov +
                                     "' and s_IOType='" + sIoType + "' and dt_CreateDate='" + insertTime + "' ";
                    if (type == "来文")
                    {
                        strSqlS += " and dt_ReceiveDate='" + dr["记录日"] + "' and ObjectType=" + InFileID;
                    }
                    else if (type == "发文")
                    {
                        strSqlS += " and dt_SendDate='" + dr["记录日"] + "' and ObjectType=" + OutFileID;
                    }
                    int nFileID = GetIDbySql(strSqlS + " order by n_FileID desc ");
                    if (nFileID > 0)
                    {
                        insertnum = nFileID;
                    }
                    if (insertnum > 0)
                    {
                        int nGovOfficeID = 0;
                        if (sClientGov.Equals("O")) //s_ClientGov C: 客户 O: 官方
                        {
                            nGovOfficeID = 21; //中国国家知识产权局
                        }
                        strSql = "SELECT COUNT(*) AS sumNum FROM dbo.T_FileInCase WHERE n_CaseID=" + hkNum +
                                 " AND n_FileID=" + insertnum;
                        int sumNumFileInCase = GetIDbySql(strSql);
                        if (sumNumFileInCase <= 0)
                        {
                            strSql = "SELECT COUNT(*) AS sumNum FROM dbo.T_FileInCase WHERE n_FileID=" + insertnum +
                                     " AND n_CaseID=" + hkNum;
                            int sumNumC = GetIDbySql(strSql);
                            if (sumNumC <= 0)
                            {
                                strSql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                                         "VALUES  (" + hkNum + " ," + insertnum + ",'Y')";
                                InsertbySql(strSql, rowid);
                            }
                        }

                        if (type == "来文")
                        {
                            strSql = "SELECT COUNT(*) AS sumNum FROM dbo.T_InFiles WHERE n_FileID=" + insertnum +
                                     " and n_GovOfficeID=" + nGovOfficeID +
                                     " and s_Distribute='Y'  and s_OFileStatus='N'";
                            int sumNum = GetIDbySql(strSql);
                            if (sumNum <= 0)
                            {
                                strSql =
                                    "INSERT INTO dbo.T_InFiles( n_FileID,n_FileCodeID,n_GovOfficeID,s_OFileStatus,s_Distribute,s_Note)" +
                                    "VALUES  (" + insertnum + ",0," + nGovOfficeID + " ,'N','Y','"+sNote+"')";
                                InsertbySql(strSql, rowid);
                            }
                        }
                        else if (type == "发文")
                        {
                            strSql =
                                "INSERT INTO dbo.T_OutFiles( n_FileID ,n_CheckedOutBy , n_GovOfficeID , s_FileStatus, dt_StatusDate ,dt_WriteDate ," +
                                "n_WriterID , n_SubmiterID ,  n_PrintNum , n_PageNum ,n_ReFileID  ,n_Count ,s_FileType ,n_LatestCheckInfoID";
                            string drValue = "VALUES  (" + insertnum + ",0 ," + nGovOfficeID + " ,'W' ,'" + DateTime.Now + "' ,'" +
                              DateTime.Now + "',0 ,0 ,1,0 ,0,0 ,'new',0";
                            if(nAgencyID>0)
                            {
                                strSql += ",n_AgencyID";
                                drValue += "," + nAgencyID;
                            }
                            strSql = strSql + ")   " + drValue + ")";
                            string sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_OutFiles WHERE n_FileID=" +
                                         insertnum;
                            int sumcount = GetIDbySql(sql);
                            if (sumcount <= 0)
                            {
                                InsertbySql(strSql, rowid);
                            }
                        }

                    }
                }
                //法律信息
                else if (dr["项目"] != null &&
                         (dr["项目"].ToString().Trim() == "PCT公开号" || dr["项目"].ToString().Trim() == "PCT公开日" ||
                          dr["项目"].ToString().Trim() == "PCT进入日" || dr["项目"].ToString().Trim() == "PCT申请号" ||
                          dr["项目"].ToString().Trim() == "PCT申请日" || dr["项目"].ToString().Trim() == "PCT办登日" ||
                          dr["项目"].ToString().Trim() == "公开号" || dr["项目"].ToString().Trim() == "公开日" ||
                          dr["项目"].ToString().Trim() == "进入实审日" || dr["项目"].ToString().Trim() == "授权公告号" ||
                          dr["项目"].ToString().Trim() == "授权公告日"))
                {
                    strSql = "UPDATE dbo.TPCase_LawInfo set " + GetColum(dr["项目"].ToString()) + "='" +
                             dr["内容"].ToString().Replace("'", "''") + "'";
                    strSql += "  WHERE n_ID IN (SELECT n_LawID FROM TPCase_Patent WHERE n_CaseID=" + hkNum + ")";
                    InsertbySql(strSql, rowid);
                }
                else if (dr["项目"] != null && dr["项目"].ToString().Trim() == "提实审日")
                {
                    strSql = "UPDATE dbo.TPCase_Patent set dt_RequestSubmitDate='" +
                             dr["内容"].ToString().Replace("'", "''") + "' WHERE n_CaseID=" + hkNum;
                    InsertbySql(strSql, rowid);
                }
                result = 1;
                #endregion
                #endregion
            }
            else
            {
                #region IDS号和IDS日

                if (dr["项目"] != null && (dr["项目"].ToString().Trim() == "IDS号" || dr["项目"].ToString().Trim() == "IDS日"))
                {
                    object obj = GetTimebySql("select s_Notes from TPCase_Patent WHERE n_CaseID=" + hkNum);
                    string notes = obj == null ? "" : obj.ToString();
                    string content = "";
                    if (!string.IsNullOrEmpty(notes))
                    {
                        content = notes + "\r\n";
                    }
                    content += Environment.NewLine + "项目：" + dr["项目"] + " 内容:" + dr["内容"];
                    string strSql = " UPDATE TPCase_Patent SET s_Notes='" + content + "' WHERE n_CaseID=" + hkNum;

                    InsertbySql(strSql, rowid);
                    //IDS号、IDS日生成一条 发官方文 记录，IDS日为发文日，IDS号为发文名称。
                    string idsName = "";
                    string time = "";
                    if (dr["项目"] != null && dr["项目"].ToString().Trim() == "IDS号")
                    {
                        idsName = dr["内容"].ToString();
                    }
                    else if (dr["项目"] != null && dr["项目"].ToString().Trim() == "IDS日")
                    {
                        time = dr["内容"].ToString();
                    }

                    //查询文件主表是否包含此文件记录
                    string strSql1 = " SELECT n_FileID FROM dbo.T_MainFiles WHERE s_Status='Y' AND s_Name='" + idsName +
                                     "' AND dt_SendDate='" + time + "' and dt_CreateDate='" + insertTime + "'";
                    int nFileID = GetIDbySql(strSql1);
                    if (nFileID > 0)
                    {
                        string strSql2 = " SELECT n_FileID FROM dbo.T_OutFiles WHERE n_FileID=" + nFileID;
                        int nFileIDOut = GetIDbySql(strSql2);
                        if (nFileIDOut > 0) //发件表和文件主表关联
                        {
                            string strSql3 = "  select n_ID FROM T_FileInCase  WHERE n_FileID =" + nFileID +
                                             " AND n_CaseID=" + hkNum;
                            int nID = GetIDbySql(strSql3);
                            if (nID <= 0) //发件和无案件关联
                            {
                                InserIntoTFileInCase(hkNum, nFileID, rowid, sNo);
                            }
                            else
                            {
                                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','已存在案件关系数据','T_FileInCase','国内-法律信息及日志表-" + rowid + "','" + strSql1.Replace("'", "''") + "')");
                            }
                        }
                        else
                        {
                            if (InserIntoTOutFiles(21, nFileID) > 0)
                            {
                                InserIntoTFileInCase(hkNum, nFileID, rowid, sNo);
                            }
                        }
                    }
                    else
                    {
                        if (InserIntoTMainFiles(idsName, insertTime, time, rowid, hkNum, sNo) > 0)
                        {
                            nFileID = GetIDbySql(strSql1);
                            if (nFileID > 0)
                            {
                                if (InserIntoTOutFiles(21, nFileID) > 0)
                                {
                                    string strSql2 = " SELECT n_FileID FROM dbo.T_OutFiles WHERE n_FileID=" + nFileID;
                                    int nFileIDOut = GetIDbySql(strSql2);
                                    if (nFileIDOut > 0) //发件表和文件主表关联
                                    {
                                        InserIntoTFileInCase(hkNum, nFileID, rowid, sNo);
                                    }
                                }
                            }
                            else
                            {
                                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','未查询到文件数据','T_MainFiles','国内-法律信息及日志表-" + rowid + "','" + strSql1.Replace("'","''") + "')");
                            }
                        }
                        else
                        {
                            InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','增加数据失败T_MainFiles','T_MainFiles','国内-法律信息及日志表-" + rowid + "')");
                        }
                    }
                }
                #endregion

                #region
                else if (dr["项目"] != null &&
                         (dr["项目"].ToString().Trim() == "办登日" || dr["项目"].ToString().Trim() == "驳回日" ||
                          dr["项目"].ToString().Trim() == "复审日"))
                {
                    //导入到官方来文，生成来文记录，内容中的日期作为来文日期
                    string Name = dr["项目"].ToString().Replace("'", "''");
                    if (Name.Equals("办登日"))
                    {
                        Name = "办登通知书";
                    }
                    else if (Name.Equals("复审日"))
                    {
                        Name = "复审通知书";
                    }
                    else if (Name.Equals("驳回日"))
                    {
                        Name = "驳回决定通知书";
                    }
                    InserIntoTMainFilesIn(hkNum, Name,
                                          dr["内容"].ToString().Replace("'", "''"), insertTime,rowid);
                }
                #endregion
            }
            return result;
        }

        #region 发文

        //文件信息基础表 官方 发文
        private int InserIntoTMainFiles(string idsName, DateTime insertTime, string time, int rowid, int hkNum, string sNo)
        {
            string strSql =
                "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,s_Abstact,ObjectType,s_SendMethod,dt_EditDate,s_Name,dt_SendDate,dt_CreateDate,s_IOType,s_ClientGov,s_Status)" +
                "VALUES  ('国外-法律信息及日志表" + rowid + "',''," + OutFileID + ",'其他','" + insertTime + "','" + idsName + "','" + time + "','" + insertTime +"','O','O','Y') ";
            int Num = InsertbySql(strSql, 0);
            if (Num <= 0)
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','增加数据失败','T_MainFiles','国内-法律信息及日志表-" + rowid + "','" + strSql.Replace("'", "''") + "')");
            }
            return Num;
        }

        //案件与发文关系 官方
        private int InserIntoTOutFiles(int nGovOfficeID, int nFileID)
        {
            string strSql =
                "INSERT INTO dbo.T_OutFiles( n_FileID ,n_CheckedOutBy , n_GovOfficeID , s_FileStatus, dt_StatusDate ,dt_WriteDate ," +
                "n_WriterID , n_SubmiterID ,  n_PrintNum , n_PageNum ,n_ReFileID  ,n_Count ,s_FileType ,n_LatestCheckInfoID)" +
                "VALUES  (" + nFileID + ",0 ," + nGovOfficeID + " ,'W' ,'" + DateTime.Now + "' ,'" + DateTime.Now +
                "',0 ,0 ,1,0 ,0,0 ,'new',0 )";
            string sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_OutFiles WHERE n_FileID=" +
                         nFileID;
            if (GetIDbySql(sql) <= 0)
            {
                return InsertbySql(strSql, 0);
            } 
            return 0;
        }

        //案件与发文关系 客户
        private int InserIntoTOutFilesCtrmst(int nFileID)
        {
            string strSql =
                "INSERT INTO dbo.T_OutFiles( n_FileID ,n_CheckedOutBy , n_GovOfficeID , s_FileStatus, dt_StatusDate ,dt_WriteDate ," +
                "n_WriterID , n_SubmiterID ,  n_PrintNum , n_PageNum ,n_ReFileID  ,n_Count ,s_FileType ,n_LatestCheckInfoID)" +
                "VALUES  (" + nFileID + ",0 ,0 ,'W' ,'" + DateTime.Now + "' ,'" + DateTime.Now +
                "',0 ,0 ,1,0 ,0,0 ,'new',0 )";
            string sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_OutFiles WHERE n_FileID=" +
                         nFileID;
            if (GetIDbySql(sql) <= 0)
            {
                return InsertbySql(strSql, 0);
            }
            return 0;
        }

        //案件与发文关系
        private void InserIntoTFileInCase(int hkNum, int nFileID, int rowid, string sNo)
        {
            string strSql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                            "VALUES  (" + hkNum + " ," + nFileID + ",'Y')";
            if (InsertbySql(strSql, 0) <= 0)
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','增加数据失败','T_FileInCase','国内-法律信息及日志表-" + rowid + "','" + strSql.Replace("'","''") + "')");
            }
        }

        #endregion

        #region 来文

        private void InserIntoTMainFilesIn(int hkNum, string project, string content, DateTime insertTime, int rowid)
        {
            string strSql =
                "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,s_Abstact,ObjectType,s_SendMethod,dt_EditDate,s_Name,dt_ReceiveDate,dt_CreateDate,s_IOType,s_ClientGov,s_Status)" +
                "VALUES  ('国外-法律信息及日志表" + rowid + "',''," + InFileID + ",'其他','" + insertTime + "','" + project + "','" + content + "','" + insertTime +
                "','I','O','Y')";
            InsertbySql(strSql, 0);
            strSql = "SELECT n_FileID FROM dbo.T_MainFiles WHERE   s_Name='" + project +
                     "' and s_ClientGov='O' and s_IOType='I' AND s_Status='Y'  and dt_ReceiveDate='" + content +
                     "' and ObjectType=" + InFileID + " and dt_CreateDate='" + insertTime + "'";
            int nFileID = GetIDbySql(strSql);
            if (nFileID > 0)
            {
                strSql = "SELECT COUNT(*) AS sumNum FROM dbo.T_FileInCase WHERE n_FileID=" + nFileID + " AND n_CaseID=" +
                         hkNum;
                int sumNumC = GetIDbySql(strSql);
                if (sumNumC <= 0)
                {
                    strSql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                             "VALUES  (" + hkNum + " ," + nFileID + ",'Y')";
                    InsertbySql(strSql, 0);
                }
                strSql = "SELECT COUNT(*) AS sumNum FROM dbo.T_InFiles WHERE n_FileID=" + nFileID +
                         " AND n_GovOfficeID=21";
                int sumNum = GetIDbySql(strSql);
                if (sumNum <= 0)
                {
                    strSql =
                        "INSERT INTO dbo.T_InFiles( n_FileID,n_FileCodeID,n_GovOfficeID,s_OFileStatus,s_Distribute)" +
                        "VALUES  (" + nFileID + ",0,21,'N','Y')";
                    InsertbySql(strSql, 0);
                } 
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + hkNum + "," + rowid + ",'','增加数据失败','T_MainFiles','国内-法律信息及日志表-" + rowid + "','" + strSql.Replace("'", "''") + "')");
            }
        }

        #endregion

        #endregion

        #region 7.国外-国外库发明人表
        private int TPCaseInventor(int rowid, DataRow dr)
        { 
            string sCountry = dr["发明人国籍"].ToString().Trim();
            dr["发明人国籍"] = GetIDbyName(sCountry, 1);

            string sNo = dr["我方卷号"].ToString().Trim();
            int hkNum = GetIDbyName(sNo, 2);
            if (hkNum.Equals(0))
            {
                //未找到“我方卷号” 
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','TPCaseInventor','国内-国外库发明人表-" + rowid + "')");
                return 0;
            }
            if (hkNum != 0)
            {
                string strSql = "SELECT n_ID FROM TPCase_Inventor WHERE n_CaseID=" + hkNum + " AND s_Name ='" + dr["发明人中文名"].ToString().Replace("'", "''") + "' AND s_NativeName='" + dr["发明人英文名"].ToString().Replace("'", "''") + "'";
                int ResultNum = GetIDbySql(strSql);
                if (ResultNum <= 0)
                {
                    int MaxSeq = GetIDbySql("SELECT TOP 1 n_Sequence  FROM TPCase_Inventor WHERE n_CaseID=" + hkNum + " ORDER BY n_Sequence DESC ");
                    strSql =
                        " INSERT INTO dbo.TPCase_Inventor(n_Sequence,n_CaseID,s_NativeName ,s_Name,n_Country,s_Address)" +
                        " VALUES(" + MaxSeq + "," + hkNum + ",'" + dr["发明人英文名"].ToString().Replace("'", "''") + "','" +
                        dr["发明人中文名"].ToString().Replace("'", "''") + "','" + dr["发明人国籍"] + "','" +
                        dr["发明人地址"].ToString().Replace("'", "''") + "')";
                }
                else
                {
                    strSql = "update TPCase_Inventor set  s_Address='" + dr["发明人地址"].ToString().Replace("'", "''")+"'";
                    if (dr["发明人国籍"] != null && !string.IsNullOrEmpty(dr["发明人国籍"].ToString()) && dr["发明人国籍"].ToString().Trim()!="0")
                    {
                        strSql += ",n_Country=" + dr["发明人国籍"].ToString();
                    }
                    strSql += " where n_ID=" + ResultNum; 
                }
                int iNum = InsertbySql(strSql, rowid);               
                if (iNum == 0)
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','TPCaseInventor导入数据错误','TPCase_Inventor','国内-国外库发明人表-" + rowid + "','" + strSql.Replace("'", "''") + "')");
                } 
                return iNum;
            }
            return 0;
        }

        #endregion

        #region 8.国外-国外库时限备注表 
        private int Case_Memo(int rowid, DataRow dr)
        {
            int result = 0;
            string sNo = dr["我方卷号"].ToString().Trim();
            int hkNum = GetIDbyName(sNo, 2);
            if (hkNum.Equals(0))
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','Case_Memo','国外-国外库时限备注表-" + rowid + "')");
                return 0;
            }
            else
            {
                if (!string.IsNullOrEmpty(dr["完成日"].ToString().Replace("'", "''")))
                {
                    result = 1;
                    string type = dr["监视类型"].ToString().Replace("'", "''");
                    string sClientGov = "C";//C: 客户  O: 官方 

                    if (!type.Equals("其它"))
                    {
                        sClientGov = "O";
                    }
                    string sd = dr["备注"] == null ? "" : "         备注:" + dr["备注"].ToString().Replace("'", "''");
                    string strSql1 =
                        " SELECT dbo.T_MainFiles.n_FileID FROM T_MainFiles LEFT JOIN dbo.T_OutFiles ON dbo.T_MainFiles.n_FileID = dbo.T_OutFiles.n_FileID" +
                        "   WHERE s_Name='监视类型：" + dr["监视类型"].ToString().Replace("'", "''") + sd + "' AND s_Status='Y' and s_Abstact='" + dr["备注"].ToString().Replace("'", "''") + " " + dr["代理人"].ToString().Replace("'", "''") + "' and s_ClientGov='" + sClientGov + "'";
                    if (dr["完成日"] != null)
                    {
                        strSql1 += " AND dt_SendDate='" + dr["完成日"].ToString().Replace("'", "''") + "'";
                    }
                    if (dr["绝限日"] != null)
                    {
                        strSql1 += " and dt_SubmitDueDate='" + dr["绝限日"].ToString().Replace("'", "''") + "'";
                    }

                    int nFileID = GetIDbySql(strSql1);
                    if (nFileID > 0)
                    {
                        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','存在此文件，无需插入','Case_Memo','国外-国外库时限备注表-" + rowid + "','" + strSql1.Replace("'", "''") + "')");
                        string strSql2 = " SELECT n_FileID FROM dbo.T_OutFiles WHERE n_FileID=" + nFileID;
                        int nFileIDOut = GetIDbySql(strSql2);
                        if (nFileIDOut > 0) //发件表和文件主表关联
                        {
                            string strSql3 = "  select n_ID FROM T_FileInCase  WHERE n_FileID =" + nFileID +
                                             " AND n_CaseID=" + hkNum;
                            int nID = GetIDbySql(strSql3);
                            if (nID <= 0) //发件和无案件关联
                            {
                                InserIntoTFileInCase(hkNum, nFileID, rowid, sNo);
                            }
                        }
                        else
                        {
                            if (InserIntoTOutFilesCtrmst(nFileID) > 0)
                            {
                                InserIntoTFileInCase(hkNum, nFileID, rowid, sNo);
                            }
                        }
                    }
                    else
                    {

                        var insertTime = DateTime.Now.AddMonths(-1);
                        string strSql =
                            "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,ObjectType,s_Status,s_SendMethod,dt_EditDate,s_Name,dt_CreateDate,s_IOType,s_ClientGov,s_Abstact ";
                        string strsql2 = "VALUES  ('国外-国外库时限备注表"+rowid+"'," + OutFileID + ",'Y','其他','" + insertTime + "','监视类型：" +
                                         dr["监视类型"].ToString().Replace("'", "''") + sd + "','" + insertTime + "','O','" + sClientGov + "','" +
                                         dr["备注"].ToString().Replace("'", "''") + " " +
                                         dr["代理人"].ToString().Replace("'", "''") + "'  ";

                        if (dr["完成日"] != null)
                        {
                            strSql += ",dt_SendDate";
                            strsql2 += ",'" + dr["完成日"].ToString().Replace("'", "''") + "'";
                        }

                        strSql += ")";
                        strsql2 += ")";
                        InsertbySql(strSql + strsql2, rowid);
                        strSql = " SELECT top 1 n_FileID FROM dbo.T_MainFiles WHERE s_Status='Y' and ObjectType=" + OutFileID +
                                 " AND  s_Name='监视类型：" + dr["监视类型"].ToString().Replace("'", "''") + sd + "' and s_Abstact='" +
                                 dr["备注"].ToString().Replace("'", "''") + " " + dr["代理人"].ToString().Replace("'", "''") +
                                 "' and s_IOType='O' and s_ClientGov='" + sClientGov + "' and dt_CreateDate='" + insertTime + "'";
                        //s_IOType I：来文 O：发文 T：其它文件 
                        //s_ClientGov  C: 客户  O: 官方 
                        if (dr["完成日"] != null)
                        {
                            strSql += " and dt_SendDate='" + dr["完成日"].ToString().Replace("'", "''") + "'";
                        }
                        nFileID = GetIDbySql(strSql + " order by n_FileID desc ");
                        if (nFileID > 0)
                        {
                            strSql =
                                "INSERT INTO dbo.T_OutFiles( n_FileID ,n_CheckedOutBy , n_GovOfficeID , s_FileStatus, dt_StatusDate ,dt_WriteDate ," +
                                "n_WriterID , n_SubmiterID ,  n_PrintNum , n_PageNum ,n_ReFileID  ,n_Count ,s_FileType ,n_LatestCheckInfoID";
                            string valueSql = "VALUES  (" + nFileID + ",0 ,21 ,'W' ,'" + insertTime + "' ,'" + insertTime +
                                              "',0 ,0 ,1,0 ,0,0 ,'new',0 ";
                            if (dr["绝限日"] != null)
                            {
                                strSql += ",dt_SubmitDueDate";
                                valueSql += ",'" + dr["绝限日"].ToString().Replace("'", "''") + "'";
                            }
                            strSql += ")";
                            valueSql += ")";
                            string sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_OutFiles WHERE n_FileID=" +
                                         nFileID;
                            if (GetIDbySql(sql) <= 0)
                            {
                                string AddSql = strSql + valueSql;
                                InsertbySql(AddSql, 0);

                            }
                            strSql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                                     "VALUES  (" + hkNum + " ," + nFileID + ",'Y')";

                            string sqlInCase = "SELECT COUNT(*) AS sumcount FROM dbo.T_FileInCase WHERE n_CaseID=" + hkNum + " and n_FileID=" + nFileID;
                            if (GetIDbySql(sqlInCase) <= 0)
                            {
                                InsertbySql(strSql, 0);
                            }
                        }
                    }
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','完成日为空，无需处理','Case_Memo','国外-国外库时限备注表-" + rowid + "')");
                }
            }
            return result;
        }

        #endregion

        #region 9.国外-实体信息表 

        private int TCaseApplicant(int rowid, DataRow dr)
        { 
            int Result = 0;
            string strSql = "";
            string sCaserial = dr["我方卷号"].ToString();
            int HkNum = GetIDbyName(sCaserial, 2);
            if (HkNum.Equals(0))
            {
                //未找到“我方卷号” 
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + HkNum + "," + rowid + ",'" + sCaserial + "','未找到“我方卷号”为:" + sCaserial + "','Case_Memo','国外-实体信息表-" + rowid + "')");
                return 0;
            }
            if (HkNum != 0)
            {
                if (dr["实体角色"] != null && dr["实体角色"].ToString().Trim() == "外代理")
                {
                    #region
                    string s_NoN = dr["实体ID"].ToString().Trim();
                    if (!string.IsNullOrEmpty(s_NoN))
                    {
                        int n_AgencyID = GetIDbyName(s_NoN, 5);
                        if (n_AgencyID > 0)
                        {
                            string strSqlA = " update TCase_Base set n_CoopAgencyToID=" + n_AgencyID +
                                             " where n_CaseID=" + HkNum;
                            InsertbySql(strSqlA, rowid);//同步代理机构ID
                            //添加代理机构联系人
                            strSql = "select n_ContactID,n_Sequence  from TCstmr_AgencyContact  where n_AgencyID=" + n_AgencyID;
                            DataTable tableA = GetDataTablebySql(strSql);
                            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(tableA.Rows.Count.ToString());
                            for (int ik = 0; ik < tableA.Rows.Count; ik++)
                            {
                                string InsertInto = " insert into TCase_ToAgencyContact(n_CaseID,n_ContactID,n_Sequence)values(" + HkNum + ",'" + tableA.Rows[ik]["n_ContactID"] + "','" + tableA.Rows[ik]["n_Sequence"] + "')";
                                InsertbySql(InsertInto, rowid);
                            }
                        }
                        else
                        {
                            InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + HkNum + "," + rowid + ",'" + sCaserial + "','实体ID未查到，“我方卷号”:" + sCaserial + "','Case_Memo','国外-实体信息表-" + rowid + "')");
                        }
                    }
                    else
                    {
                        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + HkNum + "," + rowid + ",'" + sCaserial + "','实体ID为空，“我方卷号”:" + sCaserial + "','Case_Memo','国外-实体信息表-" + rowid + "')");
                    }
                    #endregion
                }
                else if (dr["实体角色"] != null && dr["实体角色"].ToString().Trim() == "委托人")
                {
                    #region

                    if (!string.IsNullOrEmpty(dr["实体ID"].ToString().Trim().Replace("'", "''")))
                    {
                        string strSql2 = "select n_ClientID from TCstmr_Client WHERE s_ClientCode='" +
                                 dr["实体ID"].ToString().Trim().Replace("'", "''") + "'";
                        int numweituoren = GetIDbySql(strSql);
                        if (numweituoren <= 0)
                        {
                            strSql =
                                "INSERT INTO dbo.TCstmr_Client(s_Name,s_NativeName,s_State,s_City,s_Area,s_ClientCode) " +
                                "VALUES('" + dr["原名"].ToString().Trim().Replace("'", "''") + "-N','" +
                                dr["译名"].ToString().Trim().Replace("'", "''") + "','" +
                                dr["国家"].ToString().Trim().Replace("'", "''") + "','" +
                                dr["地区"].ToString().Trim().Replace("'", "''") + "','" +
                                dr["地址"].ToString().Trim().Replace("'", "''") + "','" +
                                dr["实体ID"].ToString().Trim().Replace("'", "''") + "')";
                            if (InsertbySql(strSql, rowid) > 0)
                            {
                                numweituoren = GetIDbySql(strSql2);
                            }
                            else
                            {
                                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + HkNum + "," + rowid + ",'" + sCaserial + "','增加客户信息失败','TCstmr_Client','国外-实体信息表-" + rowid + "','" + strSql.Replace("'", "''") + "')");
                            }
                        }
                        if (numweituoren > 0)
                        {
                            strSql = "  update TCase_Base set s_ClientSerial='" + dr["实体方卷号"] + "',n_ClientID=" +
                                     numweituoren + " where n_CaseID=" + HkNum;

                            int numR = InsertbySql(strSql, rowid);
                            Result = numR;
                            if (numR == 0)
                            {
                                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + HkNum + "," + rowid + ",'" + sCaserial + "','实体ID为空，“我方卷号”:" + sCaserial + "','TCase_Base','国外-实体信息表-" + rowid + "','" + strSql.Replace("'", "''") + "')");
                            }
                        }
                    }
                    #endregion
                }
                else if (dr["实体角色"] != null &&
                           (dr["实体角色"].ToString().Trim() == "申请人" || dr["实体角色"].ToString().Trim() == "申请人及委托人"))
                {
                    #region
                    string sNo = dr["实体ID"].ToString().Trim();
                    if (!string.IsNullOrEmpty(sNo))
                    {
                        int nAppID = GetIDbyName(sNo, 3);
                        int nApplicantID = 0;
                        if (nAppID > 0)
                        {
                            strSql = "SELECT n_ID FROM dbo.TCase_Applicant  WHERE n_CaseID=" + HkNum +
                                     "  AND n_ApplicantID=" + nAppID;
                            if (GetIDbySql(strSql) <= 0)
                            {
                                nApplicantID =
                                    GetIDbySql("SELECT n_AppID  FROM dbo.TCstmr_Applicant WHERE s_AppCode='" +
                                               dr["实体ID"].ToString().Trim().Replace("'", "''") + "'");
                                if (nApplicantID > 0)
                                {
                                    string sNativeName =
                                        GetTimebySql(
                                            "SELECT s_NativeName  FROM dbo.TCstmr_Applicant WHERE s_AppCode='" +
                                            dr["实体ID"].ToString().Trim().Replace("'", "''") + "'").ToString();
                                    string sName =
                                        GetTimebySql(
                                            "SELECT s_Name  FROM dbo.TCstmr_Applicant WHERE s_AppCode='" +
                                            dr["实体ID"].ToString().Trim().Replace("'", "''") + "'").ToString();
                                    string sState =
                                        GetTimebySql(
                                            "SELECT s_Area  FROM dbo.TCstmr_Applicant WHERE s_AppCode='" +
                                            dr["实体ID"].ToString().Trim().Replace("'", "''") + "'").ToString();
                                    string sStreet =
                                        GetTimebySql(
                                            "SELECT s_FirstAddress  FROM dbo.TCstmr_Applicant WHERE s_AppCode='" +
                                            dr["实体ID"].ToString().Trim().Replace("'", "''") + "'").ToString();

                                    strSql =
                                        "INSERT INTO dbo.TCase_Applicant(s_NativeName,s_Name,s_State,s_City,s_Street,s_AppSerial,n_ApplicantID,n_CaseID) " +
                                        "VALUES('" + sNativeName.Replace("'", "''") + "','" +
                                        sName.Replace("'", "''") + "','" + sState.Replace("'", "''") + "','','" +
                                        sStreet.Replace("'", "''") + "','" +
                                        dr["实体方卷号"].ToString().Trim().Replace("'", "''") + "'," + nApplicantID +
                                        "," + HkNum + ")";
                                    if (InsertbySql(strSql, rowid) <= 0)
                                    {
                                        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + HkNum + "," + rowid + ",'" + sCaserial + "','存在申请人，增加案件申请人失败','TCase_Base','国外-实体信息表-" + rowid + "','" + strSql.Replace("'", "''") + "')");
                                    }
                                }
                            }
                        }
                        else
                        {
                            strSql = "INSERT INTO dbo.TCstmr_Applicant(s_NativeName,s_Name)" +
                                     "VALUES('" + dr["原名"].ToString().Trim().Replace("'", "''") + "-N','" +
                                     dr["译名"].ToString().Trim().Replace("'", "''") + "')";
                            if (InsertbySql(strSql, rowid) > 0)
                            {
                                nApplicantID =
                                    GetIDbySql("SELECT n_AppID  FROM dbo.TCstmr_Applicant WHERE s_AppCode='" +
                                               dr["实体ID"].ToString().Trim().Replace("'", "''") + "'");
                            }
                            strSql =
                                "INSERT INTO dbo.TCase_Applicant(s_NativeName,s_Name,s_State,s_City,s_Street,s_AppSerial,n_ApplicantID,n_CaseID) " +
                                "VALUES('" + dr["原名"].ToString().Trim().Replace("'", "''") + "','" +
                                dr["译名"].ToString().Trim().Replace("'", "''") + "','" +
                                dr["国家"].ToString().Trim().Replace("'", "''") + "','" +
                                dr["地区"].ToString().Trim().Replace("'", "''") + "','" +
                                dr["地址"].ToString().Trim().Replace("'", "''") + "','" +
                                dr["实体方卷号"].ToString().Trim().Replace("'", "''") + "'," + nApplicantID + "," +
                                HkNum + ")";
                        }
                        string strSqlN = "SELECT n_ID FROM dbo.TCase_Applicant  WHERE n_CaseID=" + HkNum +
                                         "  AND n_ApplicantID=" + nAppID;
                        if (GetIDbySql(strSqlN) <= 0)
                        {
                            int numR = InsertbySql(strSql, rowid);
                            Result = numR;
                            if (numR == 0)
                            {
                                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + HkNum + "," + rowid + ",'" + sCaserial + "','不存在申请人，增加案件申请人失败','TCase_Base','国外-实体信息表-" + rowid + "','" + strSql.Replace("'", "''") + "')");
                            }
                        }

                        #region

                        if (dr["实体角色"].ToString().Trim() == "申请人及委托人")
                        {
                            string strSqAl = "select n_ClientID from TCstmr_Client WHERE s_ClientCode='" +
                                             dr["实体ID"].ToString().Trim().Replace("'", "''") + "'";
                            int nClientId = GetIDbySql(strSqAl);
                            if (nClientId > 0)
                            {
                                strSql = "  update TCase_Base set s_ClientSerial='" + dr["实体方卷号"] +
                                         "',n_ClientID=" + nClientId + " where n_CaseID=" + HkNum;
                                InsertbySql(strSql, rowid);
                            }
                            else
                            {
                                strSql =
                                    "INSERT INTO dbo.TCstmr_Client(dt_CreateDate,s_Name,s_NativeName,s_State,s_City,s_Area,s_ClientCode) " +
                                    "VALUES('" + DateTime.Now + "','" +
                                    dr["译名"].ToString().Trim().Replace("'", "''") + "-N','" +
                                    dr["原名"].ToString().Trim().Replace("'", "''") + "','" +
                                    dr["国家"].ToString().Trim().Replace("'", "''") + "','" +
                                    dr["地区"].ToString().Trim().Replace("'", "''") + "','" +
                                    dr["地址"].ToString().Trim().Replace("'", "''") + "','" +
                                    dr["实体ID"].ToString().Trim().Replace("'", "''") + "')";
                                if (InsertbySql(strSql, rowid) > 0)
                                {
                                    nClientId = GetIDbySql(strSqAl);
                                    if (nClientId > 0)
                                    {
                                        strSql = "  update TCase_Base set s_ClientSerial='" + dr["实体方卷号"] +
                                                 "',n_ClientID=" + nClientId + " where n_CaseID=" + HkNum;
                                        InsertbySql(strSql, rowid);
                                    }
                                }
                                else
                                {
                                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + HkNum + "," + rowid + ",'" + sCaserial + "','增加客户失败','TCstmr_Client','国外-实体信息表-" + rowid + "','" + strSql.Replace("'", "''") + "')");
                                }
                            }
                        }

                        #endregion
                    }
                    else
                    {
                        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + HkNum + "," + rowid + ",'" + sCaserial + "','实体ID为空，“我方卷号”:" + sCaserial + "','TCstmr_Client','国外-实体信息表-" + rowid + "','')");
                    }

                    #endregion
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + HkNum + "," + rowid + ",'" + sCaserial + "','实体ID为空，“我方卷号”:" + sCaserial + "','TCstmr_Client','国外-实体信息表-" + rowid + "','')");
                }
            }
            return Result;
        }

        #endregion
       
        #endregion
             
        #region 香港
        #region 10.香港-专利数据补充导入 
        private int HongKang(int rowid, DataRow dr)
        {
            int result = 0;
            string sNo = dr["HK卷号"].ToString().Trim();
            int hkNum = GetIDbyName(sNo, 4);
            if (hkNum.Equals(0))
            {
                //未找到“我方卷号”  
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','HongKang','香港-专利数据补充导入-" + rowid + "')");
                return 0;
            }
            if (hkNum != 0)
            {
                #region 香港表
                string sql = "SELECT * FROM TCode_BusinessType WHERE s_Code IN ('HT','HS','HD') AND n_ID IN (SELECT n_BusinessTypeID  FROM dbo.TCase_Base WHERE n_CaseID=" +
                        hkNum + " AND n_CaseID IN ( SELECT n_CaseID FROM dbo.TPCase_Patent))";
                DataTable tablew =GetDataTablebySql(sql);
                if (tablew.Rows.Count > 0)
                {
                    #region 
                    string strSql =
                        " UPDATE TCase_Base SET ObjectType=(SELECT oid FROM dbo.XPObjectType WHERE TypeName='DataEntities.Case.Patents.HongKongApplication') WHERE n_CaseID=" +
                        hkNum;
                    InsertbySql(strSql, rowid);
                    DataTable tableCount =
                        GetDataTablebySql(
                            "SELECT COUNT(*) as sumCount FROM TPCase_HongKongApplication  WHERE n_CaseID=" + hkNum);
                    string dt_1stRegisterDate = dr["第一步登记日"].ToString().Trim() == ""
                                                    ? string.Empty
                                                    : "'" + dr["第一步登记日"] + "'";
                    string dt_2ndRegisterDate = dr["第二步注册日"].ToString().Trim() == ""
                                                    ? string.Empty
                                                    : "'" + dr["第二步注册日"] + "'";
                    string dt_1stAgentReport = dr["第一步代理报告"].ToString().Trim() == ""
                                                   ? string.Empty
                                                   : "'" + dr["第一步代理报告"] + "'";
                    string dt_1stPublish = dr["第一步转公开"].ToString().Trim() == ""
                                               ? string.Empty
                                               : "'" + dr["第一步转公开"] + "'";
                    string dt_2ndAgentReport = dr["第二步代理报告"].ToString().Trim() == ""
                                                   ? string.Empty
                                                   : "'" + dr["第二步代理报告"] + "'";
                    string dt_2ndGrantReport = dr["第二步转授权"].ToString().Trim() == ""
                                                   ? string.Empty
                                                   : "'" + dr["第二步转授权"] + "'";
                    string dt_RemindShldDate = dr["维持费期限"].ToString().Trim() == ""
                                                  ? string.Empty
                                                  : "'" + dr["维持费期限"] + "'";
                    if (tableCount.Rows.Count > 0 && int.Parse(tableCount.Rows[0]["sumCount"].ToString()) > 0)
                    {
                        strSql = "UPDATE  TPCase_HongKongApplication SET s_ParentCaseSerial='" + dr["母案卷号"] +
                                 "',s_ParentCaseAppNo='" + dr["母案申请号"] + "',s_ParentCaseCountry='" + dr["母案国家"] + "'";
                        if (!string.IsNullOrEmpty(dt_1stRegisterDate))
                        {
                            strSql += ",dt_1stRegisterDate=" + dt_1stRegisterDate + "";
                        }
                        if (!string.IsNullOrEmpty(dt_2ndRegisterDate))
                        {
                            strSql += ",dt_2ndRegisterDate=" + dt_2ndRegisterDate + "";
                        }
                        if (!string.IsNullOrEmpty(dt_1stAgentReport))
                        {
                            strSql += ",dt_1stAgentReport=" + dt_1stAgentReport + "";
                        }
                        if (!string.IsNullOrEmpty(dt_1stPublish))
                        {
                            strSql += ",dt_1stPublish=" + dt_1stPublish + "";
                        }
                        if (!string.IsNullOrEmpty(dt_2ndAgentReport))
                        {
                            strSql += ",dt_2ndAgentReport=" + dt_2ndAgentReport + "";
                        }
                        if (!string.IsNullOrEmpty(dt_2ndGrantReport))
                        {
                            strSql += ",dt_2ndGrantReport=" + dt_2ndGrantReport + "";
                        }
                        if (!string.IsNullOrEmpty(dt_RemindShldDate))
                        {
                            strSql += ",dt_RemindShldDate=" + dt_RemindShldDate + "";
                        }
                        strSql += " WHERE n_CaseID=" + hkNum;
                    }
                    else
                    {
                        strSql =
                            " INSERT INTO dbo.TPCase_HongKongApplication( n_CaseID,s_ParentCaseSerial ,s_ParentCaseAppNo ,s_ParentCaseCountry";
                        string strValue = "VALUES  (" + hkNum + ",'" + dr["母案卷号"] + "','" + dr["母案申请号"] + "' ,'" +
                                          dr["母案国家"] + "'";
                        if (!string.IsNullOrEmpty(dt_1stRegisterDate))
                        {
                            strSql += ",dt_1stRegisterDate";
                            strValue += "," + dt_1stRegisterDate;
                        }
                        if (!string.IsNullOrEmpty(dt_2ndRegisterDate))
                        {
                            strSql += ",dt_2ndRegisterDate";
                            strValue += "," + dt_2ndRegisterDate;
                        }
                        if (!string.IsNullOrEmpty(dt_1stAgentReport))
                        {
                            strSql += ",dt_1stAgentReport";
                            strValue += "," + dt_1stAgentReport;
                        }
                        if (!string.IsNullOrEmpty(dt_1stPublish))
                        {
                            strSql += ",dt_1stPublish";
                            strValue += "," + dt_1stPublish;
                        }
                        if (!string.IsNullOrEmpty(dt_2ndAgentReport))
                        {
                            strSql += ",dt_2ndAgentReport";
                            strValue += "," + dt_2ndAgentReport;
                        }
                        if (!string.IsNullOrEmpty(dt_2ndGrantReport))
                        {
                            strSql += ",dt_2ndGrantReport";
                            strValue += "," + dt_2ndGrantReport;
                        }
                        strValue += ")";
                        string Content = strValue;
                        strSql += ")" + Content;
                    }
                    #endregion 
                    try
                    {
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = strSql;
                            cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                            int Insertnum = cmd.ExecuteNonQuery();
                            result = Insertnum;
                            if (Insertnum == 0)
                            { 
                                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','香港-专利数据补充导入错误','TCode_BusinessType','香港-专利数据补充导入-" + rowid + "','" + strSql.Replace("'", "''") + "')");
                            }
                        }
                    }
                    catch (Exception ex)
                    { 
                        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','香港-专利数据补充导入错误信息" + ex.Message.Replace("'","''") + "','TCode_BusinessType','香港-专利数据补充导入-" + rowid + "','" + strSql.Replace("'", "''") + "')");
                    }
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','案件不存在香港关系','TCode_BusinessType','香港-专利数据补充导入-" + rowid + "','" + sql.Replace("'", "''") + "')");
                }

                #endregion

                #region 原案信息表

                string sqlO = " SELECT n_OrigPatInfoID  FROM dbo.TPCase_Patent where n_CaseID=" + hkNum;
                int Country = GetIDbyName(dr["母案国家"].ToString(), 1);
                int InNum = GetIDbySql(sqlO);
                if (InNum > 0)
                {
                    sqlO = "update TPCase_OrigPatInfo set s_CaseSerial='" + dr["母案卷号"].ToString().Replace("'", "''") +
                           "',s_AppNo='" + dr["母案申请号"].ToString().Replace("'", "''") + "',n_OrigRegCountry=" + Country +
                           " where n_ID=" + InNum;
                }
                else
                {
                    sqlO = "INSERT INTO dbo.TPCase_OrigPatInfo(s_CaseSerial,s_AppNo,n_OrigRegCountry )" +
                           "VALUES('" + dr["母案卷号"].ToString().Replace("'", "''") + "','" +
                           dr["母案申请号"].ToString().Replace("'", "''") + "' ," + Country + ")";
                }
                if (InsertbySql(sqlO, rowid) == 0)
                { 
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','香港-专利数据补充'原案信息'导入错误信息','TPCase_OrigPatInfo','香港-专利数据补充导入-" + rowid + "','" + sqlO.Replace("'", "''") + "')");
                }
                else
                {
                    sqlO = "select n_ID from TPCase_OrigPatInfo where s_CaseSerial='" +
                           dr["母案卷号"].ToString().Replace("'", "''") + "' and s_AppNo='" +
                           dr["母案申请号"].ToString().Replace("'", "''") + "' and n_OrigRegCountry=" + Country;
                    int SQLID = GetIDbySql(sqlO);
                    if (SQLID > 0)
                    {
                        sqlO = "update dbo.TPCase_Patent set n_OrigPatInfoID=" + SQLID + " where n_CaseID=" + hkNum;
                        InsertbySql(sqlO, rowid);
                    }
                    else
                    {
                        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','原案信息不存在','TPCase_OrigPatInfo','香港-专利数据补充导入-" + rowid + "','" + sqlO.Replace("'", "''") + "')");
                    }
                }

                #endregion

                #region 案件关系

                int hkNumB = GetIDbyName(dr["母案卷号"].ToString().Trim(), 4);
                if (hkNum > 0 && hkNumB > 0)
                {
                    const string strSql = "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='香港申请' AND s_MasterName='国内母案' AND s_SlaveName='香港案' AND s_IPType='P'";
                    //案件关系配置表
                    int n_ID = GetIDbySql(strSql);
                    int NUMS =
                        GetIDbySql("SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where n_CaseIDA=" + hkNum +
                                   " and n_CaseIDB=" + hkNumB + " and n_CodeRelativeID=" + n_ID);
                    if (NUMS <= 0 && hkNum > 0)
                    {
                        InsertTCaseCaseRelative(hkNum, hkNumB, n_ID, 0);
                    }
                    else
                    {
                        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','缺少案件ID，无法建立关系','TCode_CaseRelative','香港-专利数据补充导入-" + rowid + "')");
                    }
                }

                #endregion

                #region 更新外代理人信息
                string s_CoopAgencyToNo = dr["外代理名称"].ToString().Trim();
                string strSqlAgency = "select n_AgencyID from TCstmr_CoopAgency where s_Name='" + s_CoopAgencyToNo + "' or s_NativeName='" + s_CoopAgencyToNo + "'";
                int Num = GetIDbySql(strSqlAgency);

                if (Num > 0)
                {
                    strSqlAgency = " UPDATE TCase_Base SET s_CoopAgencyToNo='" + dr["外代理文号"].ToString().Trim() + "',n_CoopAgencyToID=" + Num + " WHERE n_CaseID=" + hkNum;
                    InsertbySql(strSqlAgency, 0);
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','无外代理信息,外代理名称：" + s_CoopAgencyToNo + "','TCstmr_CoopAgency','香港-专利数据补充导入-" + rowid + "')");
                }
                #endregion

                #region 生成发文记录
                int caseNo = 0;
                int agent = 0;
                string StrSql = "select n_ClientID,n_CoopAgencyToID from tcase_base where s_caseSerial='" + sNo + "'";
                DataTable table = GetDataTablebySql(StrSql);
                if (table.Rows.Count > 0)
                {
                    caseNo = table.Rows[0]["n_ClientID"].ToString() == "" ? 0 : int.Parse(table.Rows[0]["n_ClientID"].ToString());
                    agent = table.Rows[0]["n_CoopAgencyToID"].ToString() == "" ? 0 : int.Parse(table.Rows[0]["n_CoopAgencyToID"].ToString());
                }
                //s_IOType I：来文 O：发文 T：其它文件 
                //s_ClientGov  C: 客户  O: 官方 
                string sClientGov = "";
                #region 发官方文
                //根据数据表中提供的第一步P4寄出日建立一个发外代文件，日期作为发文日；
                string One = dr["第一步P4寄出日"].ToString().Trim() == ""
                                                     ? string.Empty
                                                     : "'" + dr["第一步P4寄出日"] + "'";
                if (!string.IsNullOrEmpty(One))
                {
                    sClientGov = "O";
                    InTMainFiles("第一步P4寄出日", One, hkNum, rowid, agent, 0, sClientGov);
                }
                //根据数据表中提供的第二步P5寄出日建立一个发外代文件，日期作为发文日；
                string Two = dr["第二步P5寄出日"].ToString().Trim() == ""
                                                    ? string.Empty
                                                    : "'" + dr["第二步P5寄出日"] + "'";
                if (!string.IsNullOrEmpty(Two))
                {
                    sClientGov = "O";
                    InTMainFiles("第二步P5寄出日", Two, hkNum, rowid, agent, 0, sClientGov);
                }
                #endregion

                #region 发客户文
                string Three = dr["第一步转公开"].ToString().Trim() == ""
                                                   ? string.Empty
                                                   : "'" + dr["第一步转公开"] + "'";
                if (!string.IsNullOrEmpty(Three))
                {
                    sClientGov = "C";
                    InTMainFiles("第一步转公开", Three, hkNum, rowid, 0, caseNo, sClientGov);
                }
                string Four = dr["第二步转授权"].ToString().Trim() == ""
                                                   ? string.Empty
                                                   : "'" + dr["第二步转授权"] + "'";
                if (!string.IsNullOrEmpty(Four))
                {
                    sClientGov = "C";
                    InTMainFiles("第二步转授权", Four, hkNum, rowid, 0, caseNo, sClientGov);
                }
                string dt_stAgentReports = dr["第一步代理报告"].ToString().Trim() == ""
                                                                  ? string.Empty
                                                                  : "'" + dr["第一步代理报告"] + "'";
                if (!string.IsNullOrEmpty(dt_stAgentReports))
                {
                    sClientGov = "C";
                    InTMainFiles("第一步代理报告", dt_stAgentReports, hkNum, rowid, 0, caseNo, sClientGov);
                }
                string dt_2ndAgentReports = dr["第二步代理报告"].ToString().Trim() == ""
                                                                  ? string.Empty
                                                                  : "'" + dr["第二步代理报告"] + "'";
                if (!string.IsNullOrEmpty(dt_2ndAgentReports))
                {
                    sClientGov = "C";
                    InTMainFiles("第二步代理报告", dt_2ndAgentReports, hkNum, rowid, 0, caseNo, sClientGov);
                }
                #endregion 
                #endregion

                result = 1;
            }
            return result;
        }

        //发文主表
        private void InTMainFiles(string Name, string dtSendDate, int numHk, int i, int agent, int caseNo, string sClientGov)
        {
            DateTime insertTime = DateTime.Now;
            string strSql =
                       "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,ObjectType,s_Status,dt_EditDate,s_IOType, s_ClientGov,s_SendMethod ,s_Name ,s_Abstact,dt_CreateDate,dt_SendDate";
            string strSql2 = "VALUES  ('香港-专利数据补充导入'," + OutFileID + ",'Y','" + insertTime + "','O','" + sClientGov + "','其他','" + Name + "','" + Name + "','" + insertTime + "'," + dtSendDate;
            if (caseNo > 0)
            {
                strSql += ",n_ClientID";
                strSql2 += "," + caseNo;
            }
            strSql += ")";
            strSql2 += ")";
            string strSqlS =
                       "SELECT n_FileID FROM dbo.T_MainFiles WHERE s_Status='Y' AND s_IOType='O' and  s_ClientGov='" + sClientGov + "' and s_SendMethod='其他' and s_Name='" + Name + "'  and s_Abstact='" +
                       Name + "' and ObjectType=" + OutFileID + " and dt_SendDate=" + dtSendDate;
            if (caseNo > 0)
            {
                strSqlS += " and n_ClientID=" + caseNo;
            }
            int nFileID2 = GetIDbySql(strSqlS);
            if (nFileID2 > 0)
            {
                InTFileInCase(numHk, nFileID2, 0, agent);
            }
            else
            {
                InsertbySql(strSql + strSql2, 0);
                nFileID2 = GetIDbySql(strSqlS);
                if (nFileID2 > 0)
                {
                    InTFileInCase(numHk, nFileID2, 0, agent);
                }
            }
        }
        //案件和发文表关系
        private void InTFileInCase(int numHk, int nFileID, int i, int agent)
        {
            int fCount = 0;
            string sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_FileInCase WHERE n_FileID=" + nFileID +
                         " and n_CaseID=" + numHk;
            int sumFileInCase = GetIDbySql(sql);
            if (sumFileInCase <= 0)
            {
                sql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                      "VALUES  (" + numHk + " ," + nFileID + ",'Y')";
                int fileInCase = InsertbySql(sql, i); //记录文件、案件与程序的关系表 
                if (fileInCase <= 0)
                {
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(i + "T_FileInCase插入数据错误:" + sql);
                }
                fCount = fileInCase;
            }
            else
            {
                fCount = sumFileInCase;
            }
            if (fCount > 0)
            {
                InTOutFiles(nFileID, i, agent);
            }
        }
        //发文表
        private void InTOutFiles(int nFileID, int i, int agent)
        {
            string sqlS =
                           "INSERT INTO dbo.T_OutFiles( n_FileID ,n_CheckedOutBy  , s_FileStatus, dt_StatusDate ,dt_WriteDate ," +
                           "n_WriterID , n_SubmiterID ,  n_PrintNum , n_PageNum ,n_ReFileID  ,n_Count ,s_FileType ,n_LatestCheckInfoID";
            string strSql2 = "VALUES  (" + nFileID + ",0 ,'W' ,'" + DateTime.Now + "' ,'" +
                          DateTime.Now + "',0 ,0 ,1,0 ,0,0 ,'new',0";

            if (agent > 0)
            {
                sqlS += ",n_AgencyID";
                strSql2 += "," + agent;
            }
            sqlS += ")";
            strSql2 += ")";
            string sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_OutFiles WHERE n_FileID=" + nFileID;
            int sumcount = GetIDbySql(sql);
            if (sumcount <= 0)
            {
                int outFiles = InsertbySql(sqlS + strSql2, i);
                if (outFiles <= 0)
                {
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(i + "T_OutFiles插入数据错误:" +
                                                                               sqlS);
                }
            }
        }
         
        private void InsertTCaseCaseRelative(int caseIDA, int caseIDB, int n_ID, int s_MasterSlaveRelation)
        {
            //TCase_CaseRelative-案件关系表，其中n_CaseIDA与n_CaseIDB为两个案件，如果s_MasterSlaveRelation为1时，表示A为主案件；0表示B为主案件；-1为两者为平级，不分主从。n_CodeRelativeID为对应的关联关系配置ID
            string strSql =
                " INSERT INTO  dbo.TCase_CaseRelative ( n_CaseIDA ,  n_CaseIDB , dt_CreateDate , dt_EditDate , s_MasterSlaveRelation , n_CodeRelativeID )" +
                " VALUES  ( " + caseIDA + " , " + caseIDB + " ,  GETDATE() , GETDATE() ,  " + s_MasterSlaveRelation +
                " ,  " + n_ID + ")";
             InsertbySql(strSql, 0);
        }

        #endregion

        #region 11.优先权(香港)
        private int HongKangPriority(int rowid, DataRow dr)
        {
            int result = 0;
            int Country = GetIDbyName(dr["优先权国家"].ToString().Trim(), 1);
            int HKNum = GetIDbyName(dr["我方卷号"].ToString().Trim(), 2);
            if (HKNum.Equals(0))
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("未找到“我方卷号”为:" +
                                                                           dr["我方卷号"].ToString().
                                                                               Trim());
                return 0;
            }
            else
            {
                string strSql = "SELECT a.n_CaseID FROM tcase_Base  a  left join  TPCase_Patent b on a.n_CaseID=b.n_CaseID left join  TPCase_OrigPatInfo  c on b.n_OrigPatInfoID=c.n_ID " +
                              "where c.s_CaseSerial='" + dr["我方卷号"].ToString().Trim() + "'";

                DataTable table = GetDataTablebySql(strSql);
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    int HkNu = int.Parse(table.Rows[i]["n_CaseID"].ToString());
                    InsertTPCasePriority(HkNu, Country, dr, "香港-优先权", rowid);
                    if (dr["优先权国家"].ToString().Trim().Equals("中国")) //B为主案
                    {
                        strSql =
                            "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='国内优先权' AND s_MasterName='国内案' AND s_SlaveName='国外案' AND s_IPType='P'";
                        int n_ID = GetIDbySql(strSql);
                        InsertInto(dr["优先权号"].ToString().Trim(), HkNu, n_ID, i);
                    }
                    UpdateSeq(HkNu);
                }
            }
            return result;
        }

        private void InsertTPCasePriority(int HKNum, int Country, DataRow dr, string excelName, int rowid)
        {
            string strSql = "select n_ID from TPCase_Priority WHERE n_CaseID=" + HKNum + " AND s_PNum='" +
                                dr["优先权号"].ToString().Trim() + "'" +
                                " and n_PCountry=" + Country +
                                " and s_PDocProvided='N' and s_PTransDocProvided='N'";
            if (!string.IsNullOrEmpty(dr["优先权日"].ToString().Trim()))
            {
                strSql += " and dt_PDate='" + dr["优先权日"].ToString().Trim() + "'";
            }
            if (GetIDbySql(strSql) <= 0)
            {
                strSql =
                    "INSERT INTO dbo.TPCase_Priority( n_CaseID ,n_Sequence , n_PCountry ,s_PNum ,s_PDocProvided , s_PTransDocProvided";
                string strsql2 = "VALUES  ('" + HKNum + "',100," + Country + ",'" +
                                 dr["优先权号"].ToString().Trim() + "','N','N'";
                if (!string.IsNullOrEmpty(dr["优先权日"].ToString().Trim()))
                {
                    strSql += " , dt_PDate";
                    strsql2 += ",'" + dr["优先权日"].ToString().Trim() + "'";
                }
                strSql += ")";
                strsql2 += ")";
                string sql = strSql + strsql2;
                if (InsertbySql(sql, 0) <= 0)
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + HKNum + "," + rowid + ",'','案件已存在此优先权','TPCase_Priority','" + excelName + "-" + rowid + "','" + sql.Replace("'", "''") + "')");
                }
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + HKNum + "," + rowid + ",'','案件已存在此优先权','TPCase_Priority','" + excelName + "-" + rowid + "','" + strSql.Replace("'", "''") + "')");
            }
        }
        //重新排列优先权顺序
        private void UpdateSeq(int nCaseID)
        {
            string strSql = " SELECT  * FROM  TPCase_Priority WHERE n_CaseID=" + nCaseID + " ORDER BY dt_PDate ASC  ";
            DataTable table = GetDataTablebySql(strSql);
            for (int i = 0; i < table.Rows.Count; i++)
            {
                strSql = "update TPCase_Priority set n_Sequence=" + i + " where n_ID=" + table.Rows[i]["n_ID"];
                InsertbySql(strSql, 0);
            }
        }
        #endregion

        #endregion 

        #region 澳门
        #region 12.澳门案件-澳门延伸

        private int InsertMacaoApplication(DataRow dr, int rowid)
        {
            int Result = 0;
            string sNo=dr["EARLIER"].ToString().Trim();
            int HKNum = GetIDbyName(sNo, 2);
            int MacaoID = GetIDbyName(dr["LATER"].ToString().Trim(), 2);
            if (HKNum.Equals(0))
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + HKNum + "," + rowid + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','InsertMacaoApplication','澳门案件-澳门延伸-" + rowid + "')");
                return 0;
            }
            else
            {
                string strSql =
                    "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='澳门延伸' AND s_MasterName='国内母案' AND s_SlaveName='澳门延伸' AND s_IPType='P'";
                int n_ID = GetIDbySql(strSql);
                InsertIntoLaw(HKNum, MacaoID, n_ID, 0, "");
                Result = 1;
            }
            return Result;
        }

        #endregion
        #endregion 

        #region 公共
        #region 13.任务时限

        private int ImportTask(int rowid, DataRow dr)
        {
            int resultNum = 0;
            string sNo = dr["案件文号"].ToString().Trim();
            int hkNum = GetIDbyName(sNo, 2);
            if (hkNum.Equals(0))
            {
                //未找到“我方卷号” 
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','TCaseLaw','任务时限-" + rowid + "')");
                return resultNum;
            }
            if (hkNum != 0)
            {
                string name = dr["时限名称"].ToString();
                string time = dr["时限日期"].ToString();

                string strSql = "SELECT n_ID FROM dbo.TCode_Employee WHERE s_Name='" + dr["负责人"].ToString() + "'";
                int nID = GetIDbySql(strSql);

                //增加任务链
                string codeDeadlinegid = InsertTFCodeDeadline(name, rowid);
                //增加任务
                string taskgid = InsertTFTask(name, rowid, time, nID);
                //增加TFTaskChain
                string taskChaingid = InsertTFTaskChain(name, rowid, hkNum, sNo);
                //增加任务链节点
                InsertTFNode(taskChaingid, taskgid);
                //增加时限日期
                InsertTFDeadline(codeDeadlinegid, rowid, time, hkNum, taskChaingid);
                resultNum = 1;
            }
            return resultNum;
        }

        //任务时限名称
        private string InsertTFCodeDeadline(string Name, int i)
        {
            string gid = "";
            object obj = GetTimebySql("select g_ID from TFCode_Deadline where s_Name='" + Name + "' and s_Type=''");
            if (obj == null)
            {
                Guid guid = Guid.NewGuid();
                string strSql = "INSERT INTO  dbo.TFCode_Deadline( g_ID , s_Name,s_Type)" +
                                "VALUES  ('" + guid + "' ,'" + Name + "','' )";
                if (InsertbySql(strSql, i) > 0)
                {
                    gid = guid.ToString();
                }
            }
            else
            {
                gid = obj.ToString();
            }
            return gid;
        }

        //任务
        private string InsertTFTask(string name, int i, string time, int peopel)
        {
            string gid = "";
            Guid guid = Guid.NewGuid();
            string strSql =
                " INSERT INTO dbo.TF_Task(g_ID ,s_Name ,s_State ,s_ReadState ,dt_CreateTime ,dt_EditTime,n_ExecutorID";
            //)" +
            string strSql2 = "VALUES  ('" + guid + "' ,'旧任务-" + name + "','P' ,'R','" + DateTime.Now + "','" +
                             DateTime.Now + "'," + peopel;
            if (!string.IsNullOrEmpty(time))
            {
                strSql += ",dt_StartDate ,dt_EndDate";
                strSql2 += ",'" + time + "' ,'" + time + "'";
            }
            strSql += ")";
            strSql2 += ")";
            if (InsertbySql(strSql + strSql2, i) > 0)
            {
                gid = guid.ToString();
            }
            return gid;
        }

        private string InsertTFTaskChain(string Name, int i, int CaseID, string s_CaseSerial)
        {
            string gid = "";
            object obj =
                GetTimebySql(
                    " SELECT TOP 1 '申请号：'+s_AppNo+';案件名称:'+s_CaseName AS s_RelatedInfo1 FROM dbo.TCase_Base  WHERE n_CaseID=" +
                    CaseID);
            string sRelatedInfo2 = obj == null ? "" : obj.ToString();
            Guid guid = Guid.NewGuid();
            string strSql =
                "  INSERT INTO dbo.TF_TaskChain( g_ID ,s_Name ,s_State ,s_TriggerType ,s_RelatedObjectType ,n_RelatedObjectID ,s_RelatedInfo1 , s_RelatedInfo2 ,dt_CreateTime ,dt_EditTime)" +
                "VALUES  ('" + guid + "' ,'旧任务链-" + Name + "','P' ,'Manual' ,'Case'," + CaseID + ",'" + s_CaseSerial +
                "','" + sRelatedInfo2 + "','" + DateTime.Now + "','" + DateTime.Now + "')";
            if (InsertbySql(strSql, i) > 0)
            {
                gid = guid.ToString();
            }
            return gid;
        }

        //任务链节点
        private void InsertTFNode(string g_TaskChainGuid, string g_FormerNodeGuid)
        {
            if (!string.IsNullOrEmpty(g_FormerNodeGuid))
            {
                Guid guid = Guid.NewGuid();
                string strSql = " INSERT INTO dbo.TF_Node( g_ID ,g_TaskChainGuid ,s_Mode ,s_Type )" +
                                "VALUES  ('" + guid + "' ,'" + g_TaskChainGuid + "','N' ,'S')";
                if (InsertbySql(strSql, 0) > 0) //开始
                {
                    Guid guid2 = Guid.NewGuid();
                    strSql =
                        " INSERT INTO dbo.TF_Node( g_ID ,g_TaskChainGuid ,g_FormerNodeGuid ,s_Mode ,s_Type , g_OwnTaskGuid)" +
                        "VALUES  ('" + guid2 + "' ,'" + g_TaskChainGuid + "','" + guid + "','N' ,'T' ,'" +
                        g_FormerNodeGuid + "')";
                    if (InsertbySql(strSql, 1) > 0) //开始
                    {
                    }
                }
                else
                {
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("InsertTFNode查询错误信息:" + strSql);
                }
            } 
        }

        //时限日期
        private void InsertTFDeadline(string g_CodeDeadlineID, int i, string Time, int CaseID, string TaskChaingid)
        {
            string strSql =
                " INSERT INTO dbo.TF_Deadline(g_CodeDeadlineID ,s_RelatedObjectType ,n_RelatedObjectID ,dt_Deadline)" +
                "VALUES  ('" + g_CodeDeadlineID + "','Case'," + CaseID + ",'" + Time + "')";
            InsertbySql(strSql, 0);

            object ob = GetTimebySql(" SELECT g_ID FROM dbo.TFCode_TaskChain WHERE   s_Code='LS001'");
            string gCodeTaskChainGuid = ob == null ? "" : ob.ToString();
            Guid guid = Guid.NewGuid();
            strSql = " INSERT INTO  dbo.TFCode_DeadlineInCodeTaskChain( g_ID ,g_CodeTaskChainGuid ,g_CodeDeadlineID)" +
                     "VALUES  ( '" + guid + "','" + gCodeTaskChainGuid + "' , '" + g_CodeDeadlineID + "' )";
            if (InsertbySql(strSql, 1) > 0)
            {
                strSql = "update TF_TaskChain set g_CodeTaskChainGuid='" + gCodeTaskChainGuid + "' where g_ID='" +
                         TaskChaingid + "'";
                InsertbySql(strSql, 2);
            }
        }

        #endregion

        #region 14.部门核对
        private int UpdateOrg(DataRow row, int rowid)
        {
            int result = 0;
            string sNo = row["客户"].ToString();
            string department = row["部门"].ToString();

            //查找部门ID
            string strSql = "select n_ID from T_Department where s_Name='" + department + "'";
            int Num = GetIDbySql(strSql);
            if (Num <= 0)
            {
                string insql = "insert into T_Department(s_Name) values('" + department + "')";
                InsertbySql(insql, 0);
                Num = GetIDbySql(strSql);
            }
            //查找案件客户
            if (Num > 0)
            {
                strSql = " UPDATE TCase_Base SET n_DepartmentID=" + Num + " WHERE n_CaseID IN (select n_CaseID from tcase_base  where  right(s_caseserial,3)='" + sNo + "')";
                InsertbySql(strSql, 0);
                result = 1;
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + Num + "," + rowid + ",'" + sNo + "','未找到部门','UpdateOrg','任务时限-" + rowid + "')");
            }
            return result;
        }

        #endregion

        #region 15.申请人需要导入电子缴费单缴费人
        private int UpdateTCstmrApplicant(DataRow row, int rowid)
        {
            int result = 0;
            string sNo = row["客户号"].ToString();
            string department = row["官方收据抬头"].ToString();

            //修改申请人为电子缴费单缴费人
            string strSql = " select n_AppID from TCstmr_Applicant  where s_AppCode='" + sNo + "'";
            int Num = GetIDbySql(strSql);

            strSql = " select s_Name from TCstmr_Applicant  where s_AppCode='" + sNo + "'";
            string s_Name = GetTimebySql(strSql) != null ? GetTimebySql(strSql).ToString() : "";
            if (Num > 0)
            {
                strSql = " UPDATE TCstmr_Applicant SET s_PayFeePerson = 'Y' WHERE n_AppID=" + Num;
                InsertbySql(strSql, 0);

                //查询哪些案件包含了此申请人
                //strSql = "select  n_CaseID from TCase_Applicant  where n_ApplicantID=" + Num;
                //DataTable table = GetDataTablebySql(strSql);
                //for (int k = 0; k < table.Rows.Count; k++)
                //{
                strSql = " UPDATE TCase_Base SET s_ElecPayer='" + department.Trim() + "' WHERE n_CaseID in (select  n_CaseID from TCase_Applicant  where n_ApplicantID=" + Num+")";
                strSql += " UPDATE T_AnnualFee SET s_ElecPayer='" + s_Name.Trim() + "' WHERE n_CaseID in (select  n_CaseID from TCase_Applicant  where n_ApplicantID=" + Num + ")";
                InsertbySql(strSql, 0);
                //}
                result = 1;
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + Num + "," + rowid + ",'" + sNo + "','不存在此申请人:" + sNo + "','UpdateTCstmrApplicant','申请人需要导入电子缴费单缴费人-" + rowid + "')");
            }
            return result;
        }
        #endregion

        #region 16.相关客户-集团客户代码
        private int UpdateRelatedcustomers(DataRow row, int rowid)
        {
            int result = 0;
            string sNo = row["客户代码"].ToString();
            string relatedcustomers = row["相关客户代码"].ToString();

            if (!string.IsNullOrEmpty(sNo) && !string.IsNullOrEmpty(relatedcustomers) && relatedcustomers.Trim() != "0")
            {
                //1.查询出相关客户代码系统ID 
                string strSql = "select  n_ClientID from TCstmr_Client where s_ClientCode='" + sNo + "'";
                int Num = GetIDbySql(strSql);
                if (Num > 0)
                {
                    //1.根据客户代码查询出所有客户案件
                    strSql = "select n_CaseID from TCase_Base  where n_ClientID in (select  n_ClientID from TCstmr_Client where s_ClientCode='" + sNo + "')";
                    DataTable table = GetDataTablebySql(strSql);
                    strSql = "select n_CaseID from TCase_Applicant  where n_ApplicantID in(select n_AppID from TCstmr_Applicant  where s_AppCode='" + sNo + "')";
                    DataTable newtable = GetDataTablebySql(strSql);
                    table.Merge(newtable);
                    for (int k = 0; k < table.Rows.Count; k++)
                    {
                        InsertTCaseClients(relatedcustomers, int.Parse(table.Rows[k]["n_CaseID"].ToString()), rowid, "相关客户-集团客户代码");
                    }
                    result = 1;
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + Num + "," + rowid + ",'" + sNo + "','为查到此客户信息:" + sNo + "','UpdateTCstmrApplicant','相关客户-集团客户代码-" + rowid + "')");
                }
            }
            return result;
        }
        #endregion

        #region 17.客户要求代码导入
        private int InsertDemand(DataRow row, int i)
        {
            string sNo = row["客户ID"].ToString();
            string relatedcustomers = row["申请人ID"].ToString();
            string demand = row["要求ID"].ToString();
            string strSql = "select top 1 n_ClientID from TCstmr_Client where s_ClientCode='" + sNo + "'";
            int nClientID = GetIDbySql(strSql);
            int Result = 0;
            if (nClientID > 0)
            {
                InsertCodeDemands(demand, "Client", nClientID, 0, Result);
            }
            if (!string.IsNullOrEmpty(sNo))
            {
                //查询客户ID是否也存在申请人表内
                strSql = " select n_AppID from TCstmr_Applicant  where s_AppCode='" + sNo + "'";
                int nAppID = GetIDbySql(strSql);
                if (nAppID > 0)
                {
                    InsertCodeDemands(demand, "Applicant", 0, nAppID, Result);
                }
            }
            if (!string.IsNullOrEmpty(relatedcustomers))
            {
                //查询客户ID是否也存在申请人表内
                strSql = " select n_AppID from TCstmr_Applicant  where s_AppCode='" + relatedcustomers + "'";
                int nAppID = GetIDbySql(strSql);
                if (nAppID > 0)
                {
                    InsertCodeDemands(demand, "Applicant", 0, nAppID, Result);
                }
            }
            
            return 1;
        }

        //客户、申请人要求
        private void InsertCodeDemands(string demand, string moduleType, int nClientID, int nSourceID, int nID)
        {
            if (nID <= 0)
            {
                //根据系统要求代码查询主题和描述
                DataTable newTable = GetDataTablebySql("select n_ID,s_IPtype,s_Title,s_Description,n_DemandType from TCode_Demand WHERE s_sysdemand='" + demand + "'");
                string title = string.Empty;
                string description = string.Empty;
                string n_DemandType = string.Empty;
                string s_IPType = string.Empty;
                string n_ID = string.Empty;
                if (newTable.Rows.Count > 0)
                {
                    n_ID = newTable.Rows[0]["n_ID"].ToString();
                    n_DemandType = newTable.Rows[0]["n_DemandType"].ToString();
                    s_IPType = newTable.Rows[0]["s_IPtype"].ToString();
                    title = newTable.Rows[0]["s_Title"].ToString();
                    description = newTable.Rows[0]["s_Description"].ToString();
                }
                InsertDemand(demand, moduleType, nClientID, nSourceID, nID, n_ID, title, description, s_IPType, n_DemandType); 
            }
            else
            {
                string strSql = "update T_Demand set n_ApplicantID=" + nSourceID + ",s_ModuleType='" + moduleType + "' where n_ID=" + nID;
                InsertbySql(strSql, 0);
               
            }
        }
        private void InsertDemand(string demand, string moduleType, int nClientID, int nSourceID, int nID, string n_ID, string title, string description, string s_IPType, string n_DemandType)
        {

            DateTime time = DateTime.Now;
            //插入客户要求并且重复插入申请人要求
            string SelectSql = "select n_ID from T_Demand where  n_SysDemandID='" + n_ID + "'";//s_ModuleType='" + moduleType + "' and
            string Sql = "INSERT INTO dbo.T_Demand(s_sourcetype1,s_ModuleType,s_Title,s_Description,s_Creator,s_Editor,s_IPType,s_SysDemand,n_DemandType,dt_EditDate,dt_CreateDate,n_SysDemandID,n_CodeDemandID";
            string Sql2 = "VALUES  ('客户要求代码导入--客户/申请人','" + moduleType + "','" + title + "','" + description + "','administrator','administrator','" + s_IPType + "','" + demand + "','" + n_DemandType + "','" + time + "','" + time + "'," + n_ID + "," + n_ID;

            if (moduleType.Equals("Applicant"))//申请人
            {
                Sql += ",n_ApplicantID";
                Sql2 += "," + nSourceID;
                SelectSql += " and n_ApplicantID='" + nSourceID + "'";
            }
            else if (moduleType.Equals("Client"))//客户
            {
                Sql += ",n_ClientID";
                Sql2 += "," + nClientID;
                SelectSql += " and n_ClientID='" + nClientID + "'";
            }
            //else//客户-申请人ClientApplicant
            //{
            //    Sql += ",n_ApplicantID,n_ClientID";
            //    Sql2 += "," + nSourceID + "," + nClientID;
            //    SelectSql += " and n_ClientID='" + nClientID + "' and n_ApplicantID='" + nSourceID + "'";
            //}
            Sql += ")";
            Sql2 += ")";
            int Num = GetIDbySql(SelectSql);
            if (Num > 0)
            {
                string strSql = "update T_Demand set  s_ModuleType='" + moduleType + "' where n_ID=" + Num;
                InsertbySql(strSql, 0);
                Num = GetIDbySql(SelectSql); 
            }
            else
            {
                InsertbySql(Sql + Sql2, 0);
                GetIDbySql(SelectSql); 
            }
        }
        //案件要求Copy
        private void InsertCaseDemand(DataRow row)
        {
            string sNo = row["客户ID"].ToString();
            string relatedcustomers = row["申请人ID"].ToString();
            string demand = row["要求ID"].ToString();

            string strSql = "select  n_ClientID from TCstmr_Client where s_ClientCode='" + sNo + "'";
            int nClientID = GetIDbySql(strSql);
            strSql = "select  n_ClientID from TCstmr_Client where s_ClientCode='" + relatedcustomers + "'";
            int nClientID2 = GetIDbySql(strSql);

            strSql = "select  n_AppID from TCstmr_Applicant where s_AppCode='" + relatedcustomers + "'";
            int nSourceID = GetIDbySql(strSql);
            int nDemand = 0;
            DataTable newTable = GetDataTablebySql("select n_ID from TCode_Demand WHERE s_sysdemand='" + demand + "'");
            if (newTable.Rows.Count > 0)//要求ID
            {
                nDemand = int.Parse(newTable.Rows[0]["n_ID"].ToString());
            }

            //相关客户TCase_Clients
            strSql = "select n_CaseID,n_ClientID from TCase_Clients where n_ClientID in (" + nClientID + "," + nClientID2 + ")";
            DataTable table = GetDataTablebySql(strSql);
            for (int k = 0; k < table.Rows.Count; k++)
            {
                //if (nCaseID == 90032)
                //{
                if (SelectClient(nClientID, nClientID2, nSourceID, int.Parse(table.Rows[k]["n_CaseID"].ToString())))
                {
                    nDemand = 0;
                }
                InCase(table.Rows[k]["n_CaseID"].ToString(), "相关客户", nDemand);
                //}
            }

            //根据申请人查询增加要求
            if (nSourceID > 0)
            {
                strSql = "select  n_CaseID from TCase_Applicant  where n_ApplicantID=" + nSourceID;
                table = GetDataTablebySql(strSql);

                for (int k = 0; k < table.Rows.Count; k++)
                {
                    //if (nCaseID == 90032)
                    //{
                    if (SelectClient(nClientID, nClientID2, nSourceID, int.Parse(table.Rows[k]["n_CaseID"].ToString())))
                    {
                        nDemand = 0;
                    }
                    InCase(table.Rows[k]["n_CaseID"].ToString(), "申请人", nDemand);
                    //}
                }
            }

            //根据客户查询增加要求
            if (nClientID > 0)
            {
                strSql = "select n_CaseID from TCase_Base  where n_ClientID in (select  n_ClientID from TCstmr_Client where n_ClientID='" + nClientID + "')";
                table = GetDataTablebySql(strSql);
                for (int k = 0; k < table.Rows.Count; k++)
                {
                    //if (nCaseID == 90032)
                    //{
                    if (SelectClient(nClientID, nClientID2, nSourceID, int.Parse(table.Rows[k]["n_CaseID"].ToString())))
                    {
                        nDemand = 0;
                    }
                    InCase(table.Rows[k]["n_CaseID"].ToString(), "客户", nDemand);
                    //}
                }
            }
        }
        //如果客户和申请人都存在，则导入要求，不存在则不导入
        private bool SelectClient(int nClientID, int nClientID2, int nApplicantID, int nCaseID)
        {
            string strSql = "select n_CaseID from TCase_Clients where n_ClientID in (" + +nClientID + "," + nClientID2 + ")";
            DataTable table = GetDataTablebySql(strSql);
            strSql = "select n_CaseID from TCase_Base  where n_ClientID in (select  n_ClientID from TCstmr_Client where n_ClientID='" + nClientID + "')";
            DataTable table3 = GetDataTablebySql(strSql);
            table.Merge(table3);
            int row = table.Select("n_CaseID ='" + nCaseID + "' ").ToList().Count();


            strSql = "select  n_CaseID from TCase_Applicant  where n_ApplicantID=" + nApplicantID + " and n_CaseID=" + nCaseID;
            DataTable table2 = GetDataTablebySql(strSql);
            if (nClientID2.Equals(0) || (row > 0 && table2.Rows.Count > 0))
            {
                return true;
            }
            return false;
        }

        private void InCase(string nCaseID, string type, int nDemand)
        {

            //增加相关客户要求  
            //string strSql = "  select n_AppID from TCstmr_Applicant where s_AppCode in (select s_ClientCode from TCstmr_Client where n_ClientID in (select n_ClientID from TCase_Clients where n_CaseID=" + nCaseID + "))";
            string strSql = " select n_ClientID from TCase_Clients where n_CaseID=" + nCaseID + "";
            DataTable table = GetDataTablebySql(strSql);
            for (int k = 0; k < table.Rows.Count; k++)
            {
                int n_ClientID = int.Parse(table.Rows[k]["n_ClientID"].ToString());
                InDemand("相关客户", n_ClientID, nCaseID, nDemand);
            }
            strSql = "select n_ClientID from TCase_Base where n_CaseID=" + nCaseID;
            table = GetDataTablebySql(strSql);
            for (int k = 0; k < table.Rows.Count; k++)
            {
                int n_ClientID = int.Parse(table.Rows[k]["n_ClientID"].ToString());
                InDemand("客户", n_ClientID, nCaseID, nDemand);
            }
            strSql = "select  n_ApplicantID from TCase_Applicant  where n_CaseID=" + nCaseID;
            table = GetDataTablebySql(strSql);
            for (int k = 0; k < table.Rows.Count; k++)
            {
                int n_ClientID = int.Parse(table.Rows[k]["n_ApplicantID"].ToString());
                InDemand("申请人", n_ClientID, nCaseID, nDemand);
            }

        }
        private void InDemand(string type, int nClientID, string nCaseID, int nDemand)
        {
            string sModuleType = "Client";
            string strSql = "select n_ID from T_Demand where ";
            if (nDemand > 0)
            {
                strSql += " n_SysDemandID!=" + nDemand + " and";
            }
            if (type.Equals("相关客户"))
            {
                strSql += " n_ClientID=" + nClientID;
                sModuleType = "RelatedClient";
            }
            else if (type.Equals("客户"))
            {
                strSql += " n_ClientID=" + nClientID;
            }
            else if (type.Equals("申请人"))
            {
                strSql += " n_ApplicantID=" + nClientID;
                sModuleType = "Applicant";
            }
            DataTable newtable = GetDataTablebySql(strSql);
            for (int i = 0; i < newtable.Rows.Count; i++)
            {
                AddCase(int.Parse(newtable.Rows[i]["n_ID"].ToString()), nCaseID, sModuleType);
            }
        }

        private void AddCase(int nID, string nCaseID, string moduleType)
        {
            //根据系统要求代码查询主题和描述
            string strl = "select n_ID,s_IPtype,s_Title,s_Description,s_Creator,n_DemandType,n_SysDemandID,n_CodeDemandID,s_ModuleType,s_sysDemand from T_Demand WHERE n_ID=" + nID;

            DataTable newTable = GetDataTablebySql(strl);

            string s_IPType = string.Empty;
            string title = string.Empty;
            string description = string.Empty;
            string sCreator = string.Empty;
            string n_DemandType = string.Empty;

            string n_ID = string.Empty;
            string s_sysDemand = string.Empty;
            string n_SysDemandID = string.Empty;
            //string 
            if (newTable.Rows.Count > 0)
            {
                for (int k = 0; k < newTable.Rows.Count; k++)
                {
                    n_ID = newTable.Rows[k]["n_ID"].ToString();
                    n_DemandType = newTable.Rows[k]["n_DemandType"].ToString();
                    s_IPType = newTable.Rows[k]["s_IPtype"].ToString();
                    title = newTable.Rows[k]["s_Title"].ToString();
                    description = newTable.Rows[k]["s_Description"].ToString();
                    sCreator = newTable.Rows[k]["s_Creator"].ToString();
                    s_sysDemand = newTable.Rows[k]["s_sysDemand"].ToString();
                    n_SysDemandID = newTable.Rows[k]["n_SysDemandID"].ToString();

                    string Sql = "INSERT INTO dbo.T_Demand(s_sourcetype1,s_ModuleType,s_Title,s_Description,s_Creator,s_Editor,s_IPType,s_SysDemand,n_DemandType,dt_EditDate,dt_CreateDate,n_SysDemandID,n_CodeDemandID,s_SourceModuleType,n_SourceID,n_CaseID)" +
                                 "VALUES  ('客户要求代码-案件','Case','" + title + "','" + description + "','" + sCreator + "','" + sCreator + "','" + s_IPType + "','" + s_sysDemand + "','" + n_DemandType + "','" + DateTime.Now + "','" + DateTime.Now + "'," + n_SysDemandID + "," + n_SysDemandID + ",'" + moduleType + "'," + n_ID + ",'" + nCaseID + "')";

                    //查询是否存在此案件要求要求
                    string strSql = "select n_ID,s_SourceModuleType from T_Demand where s_ModuleType='Case'  and n_CaseID=" + nCaseID + " and  n_SysDemandID=" + n_SysDemandID;
                    DataTable Table = GetDataTablebySql(strSql);

                    //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("strSql:" + strSql + "==" + Table.Rows.Count);
                    if (Table.Rows.Count <= 0)
                    {
                        InsertbySql(Sql, 0);
                    }
                    else
                    {
                        string Type = Table.Rows[0]["s_SourceModuleType"].ToString();
                        if (moduleType.Equals("Applicant") && Type.Equals("Client"))
                        {
                            strSql = "update T_Demand set dt_EditDate='" + DateTime.Now + "', s_SourceModuleType='" + moduleType + "' where s_ModuleType='Case'  and n_CaseID=" + nCaseID + " and  n_SysDemandID=" + n_SysDemandID;
                        }
                        else if (moduleType.Equals("RelatedClient") && (Type.Equals("Applicant") || Type.Equals("Client")))
                        {
                            strSql = "update T_Demand set dt_EditDate='" + DateTime.Now + "',s_SourceModuleType='" + moduleType + "' where s_ModuleType='Case'  and n_CaseID=" + nCaseID + " and  n_SysDemandID=" + n_SysDemandID;
                        }
                        //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("Type" + Type + "====moduleType" + moduleType + "====strSql==1:" + strSql);
                        InsertbySql(strSql, 0);
                    }
                }
            }
        }

        #endregion

        #region 17-1.清除冗余客户要求_1
        private int DeleteDemand(DataRow row)
        {
            string sNo = row["客户ID"].ToString();
            string relatedcustomers = row["申请人ID"].ToString();
            string demand = row["要求ID"].ToString();

            string strSql = "select  n_ClientID from TCstmr_Client where s_ClientCode='" + sNo + "'";
            int nClientID = GetIDbySql(strSql);//客户ID

            strSql = "select  n_ClientID from TCstmr_Client where s_ClientCode='" + relatedcustomers + "'";
            int nClientID2 = GetIDbySql(strSql);//根据申请人ID查询出客户信息

            strSql = "select  n_AppID from TCstmr_Applicant where s_AppCode='" + relatedcustomers + "'";
            int nApplicantID = GetIDbySql(strSql);//申请人ID

            strSql = "select n_CaseID,n_ClientID from TCase_Clients where n_ClientID in (" + +nClientID + "," + nClientID2 + ")";
            DataTable table = GetDataTablebySql(strSql);

            strSql = "select n_CaseID,n_ClientID from TCase_Base  where n_ClientID in (select  n_ClientID from TCstmr_Client where n_ClientID='" + nClientID + "')";
            DataTable table3 = GetDataTablebySql(strSql);
            table.Merge(table3);

            for (int i = 0; i < table.Rows.Count; i++)
            {
                int nCaseID = int.Parse(table.Rows[i]["n_CaseID"].ToString());
                int nClientID_1 = int.Parse(table.Rows[i]["n_ClientID"].ToString());

                strSql = "select  n_CaseID from TCase_Applicant  where n_ApplicantID=" + nApplicantID + " and n_CaseID=" + nCaseID;
                DataTable table2 = GetDataTablebySql(strSql);
                if (table2.Rows.Count <= 0)//删除
                {
                    Delete(demand, nClientID_1, nCaseID);
                }
            }
            return 0;
        }
        private void Delete(string demand, int nClientID, int nCaseID)
        {
            //首先查询出客户+要求ID
            int nDemand = 0;
            DataTable newTable = GetDataTablebySql("select n_ID from TCode_Demand WHERE s_sysdemand='" + demand + "'");
            if (newTable.Rows.Count > 0)//要求ID
            {
                nDemand = int.Parse(newTable.Rows[0]["n_ID"].ToString());
            }
            string strSql = "select n_ID from T_Demand where n_SysDemandID=" + nDemand + " and  n_ClientID=" + nClientID;
            DataTable newtable = GetDataTablebySql(strSql);
            for (int i = 0; i < newtable.Rows.Count; i++)
            {
                DeleteCase(nCaseID, int.Parse(newtable.Rows[i]["n_ID"].ToString()), nDemand);
            }
        }
        private void DeleteCase(int nCaseID, int nID, int nDemand)
        {
            string strSql = "delete T_Demand where s_ModuleType='Case' and n_SysDemandID=" + nDemand + " and  n_CaseID=" + nCaseID + " and n_SourceID=" + nID;
            //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("清除要求:" + strSql);
            InsertbySql(strSql, 0);
        }
        #endregion

        #region 18.递交机构配置

        private int InserIntoOrgin(DataRow row, int i)
        {
            const int result = 0;

            #region 国家/地区

            string chinesecountry = row["国家"].ToString().Trim();// == "香港" ? "中国香港" : row["国家"].ToString().Trim();
            string nationalabbreviation = row["国家简码"].ToString().Trim();
            string englishcountry = row["国家（外文）"].ToString().Trim();

            string strSql = "SELECT n_ID FROM dbo.TCode_Country  WHERE  s_Name='" + chinesecountry + "'";
            int countryID = GetIDbySql(strSql);
            if (countryID <= 0)
            {
                if (!string.IsNullOrEmpty(chinesecountry))
                {
                    string strSql1 =
                        "INSERT INTO dbo.TCode_Country ( s_CountryCode , s_Name , s_OtherName , s_MadridAgreement , s_MadridProtocol , s_PCTContract , s_MultiClass , dt_CreateDate , dt_EditDate , n_FrequentNo ) VALUES  ('" +
                        nationalabbreviation + "','" + chinesecountry + "','" + englishcountry +
                        "','N','N','N' ,'N' ,GETDATE(),GETDATE(),1)";
                    if (InsertbySql(strSql1, 0) > 0)
                    {
                        countryID = GetIDbySql(strSql);
                    }
                }
            }

            #endregion

            string sRemark = "";

            #region 语种
            string[] Language = row["语种代码"].ToString().Trim().Split('/');
            string Y = "SELECT n_ID FROM dbo.TCode_Language WHERE s_LanguageCode='" + Language[0] +
                              "'";
            string[] LanguageCode = row["语种"].ToString().Trim().Split('/');
            string[] LanguageCodeEn = row["语种（外文）"].ToString().Trim().Split('/');
            int YID = GetTimebySql(Y) == null ? 0 : int.Parse(GetTimebySql(Y).ToString());
            if (YID <= 0 && !string.IsNullOrEmpty(Language[0]))//增加语种
            {
                strSql = "INSERT INTO  dbo.TCode_Language(s_OtherName, s_LanguageCode ,s_Name,dt_CreateDate ,dt_EditDate) Values('" + LanguageCodeEn[0] + "','" + Language[0] + "','" + LanguageCode[0] + "','" + DateTime.Now + "','" + DateTime.Now + "')";
                InsertbySql(strSql, 0);
                YID = GetTimebySql(Y) == null ? 0 : int.Parse(GetTimebySql(Y).ToString());
            }
            if (Language.Length > 1)
            {
                for (int j = 0; j < Language.Length; j++)
                {
                    if (Language.Length != 1 && j != 0)
                    {
                        sRemark += Language[j] + "   " + LanguageCode[j] + "   " + LanguageCodeEn[j];
                    }
                }
            }
            #endregion

            #region 官方机构
            if (countryID > 0)
            {
                string strSqlO = "SELECT  n_ID FROM dbo.TCode_Official WHERE s_Name='" +
                                 row["名称"].ToString().Trim() + "' and n_Country=" + countryID;
                DataTable table = GetDataTablebySql(strSqlO);
                if (table.Rows.Count <= 0)
                {
                    strSqlO =
                        "INSERT INTO dbo.TCode_Official( s_IPType,s_OfficialCode ,s_Name ,s_OtherName ,n_Language ,s_Phone  " +
                        ",s_Email ,n_Country ,dt_CreateDate , dt_EditDate,s_Notes )" +
                        "VALUES('P','" + row["机构编码"].ToString().Trim() + "','" + row["名称"].ToString().Trim() + "','" +
                        row["名称（外文）"].ToString().Trim().Replace("'", "''") + "'," + YID + ",'" + row["电话"].ToString().Trim() +
                        "','" + row["电子邮箱"].ToString().Trim() + "'," + countryID + ",GETDATE(),GETDATE(),'" + sRemark + "')";

                }
                else
                {
                    strSqlO = "update TCode_Official SET  s_Notes='" + sRemark + "',s_OfficialCode='" + row["机构编码"].ToString().Trim() + "',n_Language='" + YID + "',s_Email='" + row["电子邮箱"].ToString().Trim() + "',s_Phone='" + row["电话"].ToString().Trim() + "' where n_ID=" + table.Rows[0]["n_ID"].ToString();
                }
                InsertbySql(strSqlO, 0);
            }
            #endregion

            return result;
        }

        #endregion

        #region 19.转入案件 
        private int UpdateManualCreateChain(DataRow row, int rowid)
        {
            const int result = 0;
            string sNo = row["OURNO"].ToString();
            int hkNum = GetIDbyName(sNo, 2);

            if (sNo.Equals(0))
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + hkNum + "," + rowid + ",'" + sNo + "','未找到“我方卷号”为:" + sNo + "','UpdateManualCreateChain','转入案件-" + rowid + "')");
                return 0; 
            }
            else
            {
                string strSql = " UPDATE TCase_Base SET s_IsMiddleCase='Y' WHERE n_CaseID=" + hkNum;
                InsertbySql(strSql, 0);
            }
            return result;
        }
        #endregion  
      
        #region 21.相关案件-双申
        private int InsertDoubleShen(DataRow dr,int rowid)
        {
            int Result = 0;
            string sNo = dr["我方卷号1（发明）"].ToString().Trim();
            int HKNum = GetIDbyName(sNo, 2);
            int DoubleShenID = GetIDbyName(dr["我方卷号2（实用新型）"].ToString().Trim(), 2);
            if (HKNum.Equals(0))
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + HKNum + "," + rowid + ",'" + sNo + "','不存在我方卷号1（发明）:" + sNo + "','InsertDoubleShen','相关案件-双申-" + rowid + "')");
            }
            else
            {
                if (DoubleShenID.Equals(0))
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + HKNum + "," + rowid + ",'" + DoubleShenID + "','不存在我方卷号2（实用新型）:" + DoubleShenID + "','InsertDoubleShen','相关案件-双申-" + rowid + "')");
                }
                else
                {
                    string strSql =
                        "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='双申' AND s_MasterName='发明' AND s_SlaveName='实用新型' AND s_IPType='P'";
                    int n_ID = GetIDbySql(strSql);
                    InsertIntoLaw(HKNum, DoubleShenID, n_ID, 0, "");
                    Result = 1;
                }
            }
            return Result;
        }

        #endregion

        #region 22.根据业务类型修改申请方式
        private int UpdateType(DataRow row, int rowid)
        {
            const int result = 0;
            string type = row["名称"].ToString();
            if (!string.IsNullOrEmpty(type))
            {
                string strSql = "SELECT n_ID FROM TCode_BusinessType  WHERE s_Name='" + type + "'";
                int nID = GetIDbySql(strSql);

                string type1 = row["申请方式"].ToString();
                if (type1.Equals("纸件"))
                {
                    type1 = "N";
                }
                else if (type1.Equals("电子"))
                {
                    type1 = "Y";
                }
                else if (type1.Equals("不提交"))
                {
                    type1 = "U";
                }
                if (nID > 0)
                {
                    strSql = " UPDATE TCase_Base SET s_IsRegOnline='" + type1 + "' WHERE n_BusinessTypeID=" + nID;
                    InsertbySql(strSql, 0);
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + nID + "," + rowid + ",'" + type + "','业务类型为:" + type + "','UpdateType','业务类型-" + rowid + "')");
                }
            }
            return result;
        }
        private int UpdateUSAType(DataRow row, int rowid)
        {

            string sNo = row["我方文号"].ToString();
            int HkNum = GetIDbyName(sNo,2);
            if (HkNum > 0)
            {
                try
                {
                    string type = row["申请方式"].ToString();
                    if (type.ToUpper().Equals("CA申请"))
                    {
                        type = "A";
                    }
                    else if (type.ToUpper().Equals("CIP申请"))
                    {
                        type = "P";
                    }
                    string Sql = "UPDATE TCase_Base SET s_IsRegOnline='" + type + "' WHERE n_CaseID=" + HkNum;
                    return InsertbySql(Sql, 0); 
                }
                catch (Exception ex)
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + HkNum + "," + rowid + ",'" + sNo + "','更新案件状态请信息错误" + ex.Message.Replace("'", "''") + "','UpdateCasesCaseStatus','案件状态-" + rowid + "')");
                }
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + HkNum + "," + rowid + ",'" + sNo + "','不存在此案件：" + sNo + "','UpdateCasesCaseStatus','案件状态-" + rowid + "')");
            }
            return 0;
        }
        #endregion
        
        #region 20.总委托书号  
        private int UpdateTotalCommissionNumber(DataRow row, int rowid)
        {
            int result = 0;
            string sNo = row["CLIENT_NO"].ToString();

            //修改申请人为电子缴费单缴费人
            string strSql = " select n_AppID from TCstmr_Applicant  where s_AppCode='" + sNo + "'";
            int Num = GetIDbySql(strSql);

            if (Num > 0)
            {
                strSql = " UPDATE TCstmr_Applicant  ";

                strSql += row["总委号1"].ToString() == "" ? "" : "set s_HasTrustDeed='Y',s_TrustDeedNo='" + row["总委号1"].ToString() + "'";
                if (strSql.Contains("set"))
                {
                    strSql += row["英文名称"].ToString() == "" ? "" : ",s_NativeName='" + row["英文名称"].ToString() + "'";
                }
                else
                {
                    strSql += row["英文名称"].ToString() == "" ? "" : " set s_NativeName='" + row["英文名称"].ToString() + "'";
                }
                if (strSql.Contains("set"))
                {
                    strSql += row["中文名称1"].ToString() == "" ? "" : ",s_Name='" + row["中文名称1"].ToString() + "'";
                }
                else
                {
                    strSql += row["中文名称1"].ToString() == "" ? "" : " set s_Name='" + row["中文名称1"].ToString() + "'";
                }
                strSql += "  WHERE n_AppID=" + Num;
                strSql += "   update TCase_Applicant set  s_TrustDeedNo='" + row["总委号1"].ToString() + "' where n_ApplicantID=" + Num;

                //增加译名
                if (!string.IsNullOrEmpty(row["中文名称2"].ToString()) || !string.IsNullOrEmpty(row["总委号2"].ToString()))
                {
                    strSql +=
                        " INSERT INTO dbo.TCstmr_AppTransLatedName( s_AppTransLatedName ,s_TrustdeedNum ,s_AppTransLatedNameUse ,n_AppID)VALUES  ('" +
                        row["中文名称2"].ToString() + "','" + row["总委号2"].ToString() + "','P, T, C, D, O'," + Num + ")";
                }
                if (!string.IsNullOrEmpty(row["中文名称3"].ToString()) || !string.IsNullOrEmpty(row["总委号3"].ToString()))
                {
                    strSql +=
                        " INSERT INTO dbo.TCstmr_AppTransLatedName( s_AppTransLatedName ,s_TrustdeedNum ,s_AppTransLatedNameUse ,n_AppID)VALUES  ('" +
                        row["中文名称3"].ToString() + "','" + row["总委号3"].ToString() + "','P, T, C, D, O'," + Num + ")";
                }
                InsertbySql(strSql, 0);

                result = 1;
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + Num + "," + rowid + ",'" + sNo + "','不存在申请人:" + sNo + "','UpdateTotalCommissionNumber','总委托书号-" + rowid + "')");
            }
            return result;
        }
        #endregion 
        #region 注释
     
        #region 23.案件状态 注释
        //private int UpdateCasesCaseStatus(DataRow row, int rowid)
        //{
        //    int result = 0;
        //    string sNo = row["申请号"].ToString();
        //    if (!string.IsNullOrEmpty(sNo))
        //    {
        //        int HkNum = GetIDbyName(sNo.Substring(2), 7);
        //        if (HkNum > 0)
        //        {
        //            try
        //            {
        //                string Sql = "UPDATE TCase_Base SET s_CaseStatus='" + row["法律状态"].ToString() + "' WHERE n_CaseID=" + HkNum;
        //                InsertbySql(Sql, 0);
        //                result = 1;
        //            }
        //            catch (Exception ex)
        //            { 
        //                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + HkNum + "," + rowid + ",'" + sNo + "','更新案件状态请信息错误" + ex.Message.Replace("'", "''") + "','UpdateCasesCaseStatus','案件状态-" + rowid + "')");
        //            }
        //        }
        //        else
        //        {
        //            InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(" + HkNum + "," + rowid + ",'" + sNo + "','不存在此案件：" + sNo + "','UpdateCasesCaseStatus','案件状态-" + rowid + "')");
        //        }
        //    }
        //    return result;
        //}
        #endregion

        #endregion 

        #region 24.客户要求
        private int InsertDemandClient(DataRow dr, int rowid)
        {
            int nID = InsertDemandType(dr["标题"].ToString());
            string sClientCode = "0";
            if (dr["客户编号"] != null && !string.IsNullOrEmpty(dr["客户编号"].ToString()))
            {
                sClientCode = dr["客户编号"].ToString();
                int nApplicantID =
                     GetIDbySql("SELECT n_ClientID FROM  dbo.TCstmr_Client   WHERE s_ClientCode='" + sClientCode + "'");
                if (nApplicantID > 0)
                {
                    InsertUserDemandClient(nID, dr["收到方式"].ToString(), dr["指示人"].ToString(), dr["标题"].ToString(),
                                     dr["描述"].ToString(), dr["收到日"].ToString(), "客户", sClientCode, rowid, nApplicantID);
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(0," + rowid + ",'" + sClientCode + "','不存在客户：" + sClientCode + "','InsertDemandClient','客户要求整理Client-客户-" + rowid + "')");
                }
            }
            if (dr["申请人编号"] != null && !string.IsNullOrEmpty(dr["申请人编号"].ToString()))
            {
                sClientCode = dr["申请人编号"].ToString();
                int nClientID =
                    GetIDbySql("SELECT n_AppID FROM  dbo.TCstmr_Applicant  WHERE s_AppCode='" + sClientCode + "'");
                if (nClientID > 0)
                {
                    InsertUserDemandClient(nID, dr["收到方式"].ToString(), dr["指示人"].ToString(), dr["标题"].ToString(),
                                     dr["描述"].ToString(), dr["收到日"].ToString(), "申请人", sClientCode, rowid, nClientID);
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName) VALUES(0," + rowid + ",'" + sClientCode + "','不存在申请人：" + sClientCode + "','InsertDemandClient','客户要求整理Client-申请人-" + rowid + "')");
                }
            }
            return 1;
        }
        //要求标题
        private int InsertDemandType(string sCodeDemandType)
        {
            string strSql = "select n_ID from TFCode_DemandType where s_Name='" + sCodeDemandType + "'";
            string strSql2 = "select n_ID from TCode_Demand where s_Title='" + sCodeDemandType + "' and s_IPType='P'";
            int num = GetIDbySql(strSql);

            if (num <= 0 && GetIDbySql(strSql2) <= 0)
            {
                string InsertSql = "INSERT INTO dbo.TFCode_DemandType( s_Name)"
                                   + "VALUES  ( '" + sCodeDemandType + "')";
                if (InsertbySql(InsertSql, 0) > 0)
                {
                    num = GetIDbySql(strSql);
                    if (GetIDbySql(strSql2) <= 0)
                    {
                        InsertSql = "INSERT INTO dbo.TCode_Demand ( s_IPType , s_Type , s_Title , n_IsActive , dt_CreateDate , dt_EditDate ,  n_DemandType )"
                                    + "VALUES  ( 'P' ,'" + sCodeDemandType + "' ,'" + sCodeDemandType +
                                    "' ,1 ,GETDATE(),GETDATE()," + num + ")";
                        InsertbySql(InsertSql, 0);
                    }
                }
            }
            else
            {
                if (GetIDbySql(strSql2) <= 0)
                {
                    string InsertSql = "INSERT INTO dbo.TCode_Demand ( s_IPType , s_Type , s_Title , n_IsActive , dt_CreateDate , dt_EditDate ,  n_DemandType )"
                                       + "VALUES  ( 'P' ,'" + sCodeDemandType + "' ,'" + sCodeDemandType +
                                       "' ,1 ,GETDATE(),GETDATE()," + num + ")";
                    InsertbySql(InsertSql, 0);
                }
            }
            return num;
        }

        private void InsertUserDemandClient(int nID, string sReceiptMethod, string sAssignor, string sTitle,
                                      string sDescription, string dtReceiptDate, string type, string sClientCode, int rowid, int sClientCodeID)
        {
            string strSql =
                "INSERT INTO dbo.T_Demand(s_sourcetype1,n_DemandType,dt_CreateDate,dt_EditDate,s_ReceiptMethod,s_Assignor,s_Title,s_Description";
            string strSql1 = " VALUES  ('客户要求-" + type + "-" + rowid + "'," + nID + ",'" + DateTime.Now + "','" + DateTime.Now + "','" +
                             sReceiptMethod.Replace("'", "''") + "','" + sAssignor.Replace("'", "''") + "','" +
                             sTitle.Replace("'", "''") + "','" + sDescription.Replace("'", "''") + "'";
            if (!string.IsNullOrEmpty(dtReceiptDate))
            {
                strSql += ",dt_ReceiptDate";
                strSql1 += ",'" + dtReceiptDate + "'";
            }
            if (type.Equals("申请人"))
            {
                strSql += ",s_ModuleType,n_ApplicantID";
                strSql1 += ",'Applicant'," + sClientCodeID; 
            }
            else if (type.Equals("客户"))
            {
                strSql += ",s_ModuleType,n_ClientID";
                strSql1 += ",'Client'," + sClientCodeID;
            }

            strSql += ")";
            strSql1 += ")";
            string sql=strSql + strSql1;
            if (InsertbySql(strSql + strSql1, 0) <= 0)
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(0," + rowid + ",'" + sClientCode + "','插入要求报错','InsertUserDemandClient','客户要求整理Client-" + rowid + "','" + sql.Replace("'", "''") + "')");
            }
        }

        #endregion
        #endregion

        #region 按钮事件
        #region 1.更新数据
        private void button2_Click(object sender, EventArgs e)
        {
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============更新数据按钮 Start =================");
            sumCount = 0;
            conn.Open();
            conn.ChangeDatabase(com_databasename.Text.Trim()); //重新指定数据库 
            DateTime begin = DateTime.Now;

            #region
            string strSql = " update TCstmr_Applicant  set n_Country=(select n_Country from TCstmr_Client where TCstmr_Client.n_ClientID=TCstmr_Applicant.n_ClientID)";
            InsertbySql(strSql, 0);

            strSql = "SELECT n_CaseID FROM dbo.TCase_Base";
            DataTable table = GetDataTablebySql(strSql);

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
                DataTable tableApplicant = GetDataTablebySql(strSql);
                if (tableApplicant.Rows.Count > 0)
                {
                    string nCountry = tableApplicant.Rows[0]["n_Country"].ToString();
                    strSql = " UPDATE TCase_Base SET n_AppCountry=" + nCountry + " WHERE n_CaseID=" + table.Rows[i]["n_CaseID"];
                    InsertbySql(strSql, i); //申请人国家需要根据申请人代表自动带入 
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
            InsertbySql(strSql, 0);

            //根据是否填写了申请号更新案件提交状态
            strSql = @"
UPDATE dbo.TCase_Base SET s_SubmitStatus = 'Y'
FROM dbo.TPCase_Patent
INNER JOIN dbo.TPCase_LawInfo ON dbo.TPCase_LawInfo.n_ID = dbo.TPCase_Patent.n_LawID
WHERE dbo.TCase_Base.n_CaseID = dbo.TPCase_Patent.n_CaseID AND ISNULL(dbo.TPCase_LawInfo.s_AppNo,'') <> ''
";
            InsertbySql(strSql, 0);

            strSql = @"
UPDATE dbo.TCase_Base SET s_SubmitStatus = 'N'
FROM dbo.TPCase_Patent
INNER JOIN dbo.TPCase_LawInfo ON dbo.TPCase_LawInfo.n_ID = dbo.TPCase_Patent.n_LawID
WHERE dbo.TCase_Base.n_CaseID = dbo.TPCase_Patent.n_CaseID AND ISNULL(dbo.TPCase_LawInfo.s_AppNo,'') = ''
";
            InsertbySql(strSql, 0);

            //根据注册国家更新案件递交机构
            strSql = @"
UPDATE dbo.TCase_Base SET n_OfficeID = dbo.TCode_Official.n_ID
FROM dbo.TCode_Official
WHERE dbo.TCase_Base.n_RegCountry = dbo.TCode_Official.n_Country AND dbo.TCode_Official.s_IPType LIKE '%P%'
";
            InsertbySql(strSql, 0);

            //PCT国际案件：PCT国际申请没有填写PCT申请号，只填写了申请号
            strSql = "update TPCase_LawInfo set s_PCTAppNo=s_AppNo    where n_ID in (select n_LawID from TPCase_Patent where n_CaseID in (select n_CaseID from tcase_base where n_BusinessTypeID in (select n_ID from tcode_BusinessType where s_Name='PCT国际阶段申请')))  and s_AppNo  is not null  and isnull(s_PCTAppNo,'')=''";
            InsertbySql(strSql, 0);

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
            InsertbySql(strSql, 0);

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
            InsertbySql(strSql, 0);

            DateTime end = DateTime.Now;
            TimeSpan ts = end.Subtract(begin).Duration();
            xfrmProcess.progressBarControl.Position = xfrmProcess.progressBarControl.Properties.Maximum;
            xfrmProcess.lbTotalTime.Text = ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒";
            xfrmProcess.lbTotalSum.Text = num + "条";
            Application.DoEvents();
            Thread.Sleep(8000);
            xfrmProcess.Focus();
            xfrmProcess.Close();
            conn.Close();
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("============="+ num + "条"+ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒"+"更新数据按钮 End =================");
        }
        #endregion

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
                GetDataTablebySql(
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
WHERE n_BusinessTypeID IN (SELECT n_ID FROM dbo.TCode_BusinessType WHERE s_Name='PCT国际阶段申请') AND ISNULL(dbo.TPCase_LawInfo.s_PCTAppNo,'') <> ''");
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
                            GetDataTablebySql(
                                string.Format(
@"SELECT DISTINCT dbo.TCase_Base.n_CaseID
FROM dbo.TCase_Base
INNER JOIN dbo.TPCase_Patent ON dbo.TPCase_Patent.n_CaseID = dbo.TCase_Base.n_CaseID
INNER JOIN dbo.TPCase_LawInfo ON dbo.TPCase_LawInfo.n_ID = dbo.TPCase_Patent.n_LawID
WHERE n_BusinessTypeID NOT IN (SELECT n_ID FROM dbo.TCode_BusinessType WHERE s_Name='PCT国际阶段申请') 
AND dbo.TPCase_LawInfo.s_PCTAppNo = '{0}'", sAppNo));

                        var listPCTNCaseID = tablePCTN.Rows.Cast<DataRow>().Select(dr => dr["n_CaseID"].ToString()).ToList();
                        var listPCTNCaseIDCopy = tablePCTN.Rows.Cast<DataRow>().Select(dr => dr["n_CaseID"].ToString()).ToList();
                        string strSql =
"SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='PCT国家' AND s_MasterName='同PCT案' AND s_SlaveName='同PCT案' AND s_IPType='P'";
                        int n_ID = GetIDbySql(strSql);
                        string strSqlPct =
    "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='PCT' AND s_MasterName='PCT国际案' AND s_SlaveName='PCT国家案' AND s_IPType='P'";
                        int nIDPCT = GetIDbySql(strSqlPct);

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
            xfrmProcess.lbTotalSum.Text = num + "条";
            Application.DoEvents();
            Thread.Sleep(8000);
            xfrmProcess.Focus();
            xfrmProcess.Close();
            conn.Close();
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============" +num + "条"+ ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒" + "=================");
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============同步PCT案件关系 End =================");
        }

        //增加案件关系表
        private void InsertIntoLaw(int nCaseID, int HKNum, int n_ID, int i, string type)
        {
            int NUMS =
                GetIDbySql("SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where n_CaseIDA=" + HKNum +
                           " and n_CaseIDB=" + nCaseID + " and n_CodeRelativeID=" + n_ID);
            if (!string.IsNullOrEmpty(type)) //同族
            {
                NUMS =
                    GetIDbySql("SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where ((n_CaseIDA=" + HKNum +
                               " and n_CaseIDB=" + nCaseID + ") or (n_CaseIDA=" + nCaseID + " and n_CaseIDB=" + HKNum +
                               ") )and n_CodeRelativeID=" + n_ID);
            }

            if (NUMS <= 0 && HKNum > 0)
            {
                InsertTCaseCaseRelative(HKNum, nCaseID, n_ID, 0);
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
            DataTable table = GetDataTablebySql(strSql);

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
                DataTable relationtable = GetDataTablebySql(strSql);

                //3、将所有案件与这个美国案建立起IDS关系，注意要在集合中去除该案件本身。 
                DataView dv = new DataView(relationtable);
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
            xfrmProcess.lbTotalSum.Text = num + "条";
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
            string strSql =
                    "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='IDS' AND s_MasterName='递交方' AND s_SlaveName='接收方' AND s_IPType='P'";
            int n_ID = GetIDbySql(strSql);
            if (caseID > 0)
            {
                int NUMS =
                    GetIDbySql("SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where ((n_CaseIDA=" + HKNum +
                               " and n_CaseIDB=" + caseID + ") or  (n_CaseIDA=" + caseID + " and n_CaseIDB=" + HKNum + "))and n_CodeRelativeID=" + n_ID);
                if (NUMS <= 0 && HKNum > 0)
                {
                    InsertTCaseCaseRelative(HKNum, caseID, n_ID, 0);
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
            DataTable table = GetDataTablebySql(strSql);

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
            DataTable tablePCT = GetDataTablebySql(strSql);

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
            xfrmProcess.lbTotalSum.Text = num + "条";
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
            string strSql =
                    "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='同族专利' AND s_MasterName='同族' AND s_SlaveName='同族' AND s_IPType='P'";
            int n_ID = GetIDbySql(strSql);
            if (caseID > 0)
            {
                int NUMS =
                    GetIDbySql("SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where ((n_CaseIDA=" + HKNum +
                               " and n_CaseIDB=" + caseID + ") or  (n_CaseIDA=" + caseID + " and n_CaseIDB=" + HKNum + "))and n_CodeRelativeID=" + n_ID);
                if (NUMS <= 0 && HKNum > 0)
                {
                    InsertTCaseCaseRelative(HKNum, caseID, n_ID, 0);
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
            string strSql = "SELECT n_CaseID FROM dbo.TCase_Base WHERE n_CaseID IN (SELECT DISTINCT n_CaseID from TPCase_Priority  group by n_CaseID having count(1) >= 2)";
            DataTable table = GetDataTablebySql(strSql);
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
            xfrmProcess.lbTotalSum.Text = num + "条";
            Application.DoEvents();
            Thread.Sleep(8000);
            xfrmProcess.Focus();
            xfrmProcess.Close();
            conn.Close();
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============" + num + "条" + ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒" + "=================");
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============更新优先权顺序 End =================");
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


            string strSql = "SELECT DISTINCT n_CaseIDB FROM TCase_CaseRelative WHERE  s_MasterSlaveRelation=0 AND n_CodeRelativeID=4";
            DataTable table = GetDataTablebySql(strSql);
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
                    //if (nCaseIDA.Equals(70871))
                    //{
                    strSql = "SELECT n_CaseIDA FROM TCase_CaseRelative WHERE  s_MasterSlaveRelation=0 AND n_CodeRelativeID=4 and n_CaseIDB=" + nCaseIDA;
                    DataTable newtable = GetDataTablebySql(strSql);
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
                    //}
                }
            }


            DateTime end = DateTime.Now;
            TimeSpan ts = end.Subtract(begin).Duration();
            xfrmProcess.progressBarControl.Position = xfrmProcess.progressBarControl.Properties.Maximum;
            xfrmProcess.lbTotalTime.Text = ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒";
            xfrmProcess.lbTotalSum.Text = num + "条";
            Application.DoEvents();
            Thread.Sleep(8000);
            xfrmProcess.Focus();
            xfrmProcess.Close();
            conn.Close();
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============" + num + "条" + ts.Days + "天" + ts.Hours + "小时" + ts.Minutes + "分" + ts.Seconds + "秒" + "=================");
            Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("=============更新案件相关关系 End =================");
        }
        private void InsertTCaseCaseRelative(int caseIDA, int caseIDB)
        {
            string strSql =
                     "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='相关案件' AND s_MasterName='' AND s_SlaveName='' AND s_IPType='P'";

            int n_ID = GetIDbySql(strSql);
            if (n_ID > 0)
            {
                strSql =
                   "SELECT count(*) as Sunm FROM dbo.TCase_CaseRelative WHERE ((n_CaseIDA=" + caseIDA + " AND n_CaseIDB=" + caseIDB + ") OR (n_CaseIDA=" + caseIDB + " AND n_CaseIDB=" + caseIDA + ") ) AND s_MasterSlaveRelation=0 AND n_CodeRelativeID=" + n_ID;

                int sum = GetIDbySql(strSql);

                if (sum <= 0)
                {
                    strSql =
                        " INSERT INTO  dbo.TCase_CaseRelative ( n_CaseIDA ,  n_CaseIDB , dt_CreateDate , dt_EditDate , s_MasterSlaveRelation , n_CodeRelativeID )" +
                        " VALUES  ( " + caseIDB + " , " + caseIDA + " ,  GETDATE() , GETDATE() , 0,  " + n_ID + ")";
                    InsertbySql(strSql, 0);
                }
            }
        }
        #endregion

        #endregion 

        #endregion

        #region 重复调用方法

        #region 实体信息

        //增加用户与案件相关联  【相关客户】
        private void InsertTCaseClients(string s_ClientCode, int n_CaseID, int i, string tableName)
        {
            int Result = 0;
            //判断是否存在此人
            string strSql = "SELECT count(*) sunmCount FROM TCase_Clients WHERE n_CaseID=" + n_CaseID +
                            " AND n_ClientID=" + s_ClientCode;
            if (GetIDbySql(strSql) <= 0)
            {
                //查询人员ID
                string Sql = " SELECT n_ClientID FROM TCstmr_Client WHERE s_ClientCode='" + s_ClientCode + "'";
                int n_ClientID = 0;
                object obj = GetTimebySql(Sql);
                if (obj != null)
                {
                    n_ClientID = int.Parse(GetTimebySql(Sql).ToString());
                }
                if (n_ClientID > 0)
                {
                    Sql = "INSERT INTO dbo.TCase_Clients( n_CaseID, n_ClientID)" +
                          "VALUES  (" + n_CaseID + " , " + n_ClientID + ")";
                    Result = InsertbySql(Sql, i);
                    if (Result <= 0)
                    {
                        InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + n_CaseID + "," + i + ",'用户编号：" + s_ClientCode + "','增加用户与案件关系失败[InsertTCaseClients]','TCase_Clients','" + tableName + "-" + i + "','" + Sql.Replace("'", "''") + "')");
                        Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("增加用户与案件关系失败[InsertTCaseClients]-:" + tableName + "==strSql:" + Sql.Replace("'", "''"));
                    }
                }
                else
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + n_CaseID + "," + i + ",'用户编号：" + s_ClientCode + "','未查询到用户信息','TCstmr_Client','" + tableName + "-" + i + "','" + Sql.Replace("'", "''") + "')");
                }
            }
            else
            {
                InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,ExcelName,str2) VALUES(" + n_CaseID + "," + i + ",'用户编号：" + s_ClientCode + "','案件未存在此用户的相关客户信息','TCase_Clients','" + tableName + "-" + i + "','" + strSql.Replace("'", "''") + "')");
            }
        }

        //增加案件处理人信息
        private void InsertUser(string s_InternalCode, int n_CaseID, int i, string NameType, string tableName)
        {
            //查询人员ID
            string Sql = "SELECT n_ID FROM dbo.TCode_Employee WHERE s_InternalCode='" + s_InternalCode +
                         "'  OR s_Name='" + s_InternalCode + "'";
            int n_AssignorID = GetIDbySql(Sql);

            //案件角色
            string Sql2 = "SELECT n_ID FROM dbo.TCode_CaseRole WHERE s_Name LIKE '%" + NameType + "%'";
            int n_CaseRoleID = GetIDbySql(Sql2);

            if (n_AssignorID > 0)
            {
                Sql = "SELECT COUNT(*) AS sumCount FROM dbo.TCase_Attorney WHERE n_CaseID=" + n_CaseID +
                      " AND n_AttorneyID=" + n_AssignorID + " and n_CaseRoleID=" + n_CaseRoleID;
                if (GetIDbySql(Sql) <= 0)
                {
                    Sql =
                        "INSERT INTO dbo.TCase_Attorney( n_CaseID ,dt_AssignDate,n_AssignorID,n_AttorneyID,n_CaseRoleID)" +
                        "VALUES  (" + n_CaseID + " , '" + DateTime.Now + "',1000131," + n_AssignorID + "," +
                        n_CaseRoleID + ")";
                    InsertbySql(Sql, i);
                }
            }
            else
            {
                if (!NameType.Equals("代理部-新申请阶段-办案人"))
                {
                    InsertbySql("INSERT INTO [dbo].[Table_1]([n_CaseID],[n_RowID],[s_CaseSerial],strsql,tableName,str2,ExcelName) VALUES(" + n_CaseID + "," + i + ",'" + NameType + "','" + Sql.Replace("'", "''") + "','TCode_Employee','" + Sql.Replace("'", "''") + "','" + tableName + "')");
                }
                //Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(s_InternalCode + "信息未找到");
            }
        }
        #endregion

        #region 公共信息

        /// <summary>
        /// 查询我方卷号  国家
        /// </summary>
        /// <param name="rowName">名称</param>
        /// <param name="type">1:国家  2.案件</param>
        /// <returns></returns>
        private int GetIDbyName(string rowName, int type)
        {
            string strSql = "";
            //type 1:国家  2.案件
            int retName = 0;
            var sqlColumns = new List<string>();
            if (type == 1)
            {
                #region 查找国家ID

                try
                {
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        if (rowName.Trim().Equals("欧洲"))
                        {
                            rowName = "欧洲专利局";
                        }
                        strSql = @"select n_ID from TCode_Country where s_CountryCode='" + rowName + "' OR s_Name='" +
                                 rowName + "'";
                        cmd.CommandText = strSql;
                        cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                        using (SqlDataReader reader = cmd.ExecuteReader())
                            while (reader.Read()) sqlColumns.Add(reader[0].ToString());
                    }
                }
                catch (Exception ex)
                {
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("1.GetIDbyName导入错误信息:" + ex.Message +
                                                                               "==SQL:" + strSql);
                    //MessageBox.Show(ex.Message);
                }

                #endregion
            }
            else if (type == 2)
            {
                #region 查找提案ID

                try
                {
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        strSql = @"select n_CaseID from TCase_Base where s_CaseSerial='" + rowName + "'";
                        cmd.CommandText = @"select n_CaseID from TCase_Base where s_CaseSerial='" + rowName + "'";
                        cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                        using (SqlDataReader reader = cmd.ExecuteReader())
                            while (reader.Read()) sqlColumns.Add(reader[0].ToString());
                    }
                }
                catch (Exception ex)
                {

                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("2.GetIDbyName导入错误信息:" + ex.Message +
                                                                               "==SQL:" + strSql);
                }

                #endregion
            }
            else if (type == 3)
            {
                #region 查找客户信息

                try
                {
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        strSql = @"SELECT n_AppID FROM TCstmr_Applicant WHERE s_AppCode='" + rowName + "'";
                        cmd.CommandText = @"SELECT n_AppID FROM TCstmr_Applicant WHERE s_AppCode='" + rowName + "'";
                        cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                        using (SqlDataReader reader = cmd.ExecuteReader())
                            while (reader.Read()) sqlColumns.Add(reader[0].ToString());
                    }
                }
                catch (Exception ex)
                {
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("3.GetIDbyName导入错误信息:" + ex.Message +
                                                                               "==SQL:" + strSql);
                }

                #endregion
            }
            else if (type == 4)
            {
                #region 查找提案转为专利

                try
                {
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        strSql =
                            @"SELECT n_CaseID FROM dbo.TPCase_Patent WHERE n_CaseID IN (SELECT n_CaseID from TCase_Base where s_CaseSerial='" +
                            rowName + "')";
                        cmd.CommandText =
                            @"SELECT n_CaseID FROM dbo.TPCase_Patent WHERE n_CaseID IN (SELECT n_CaseID from TCase_Base where s_CaseSerial='" +
                            rowName + "')";
                        cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                        using (SqlDataReader reader = cmd.ExecuteReader())
                            while (reader.Read()) sqlColumns.Add(reader[0].ToString());
                    }
                }
                catch (Exception ex)
                {
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("4.GetIDbyName导入错误信息:" + ex.Message +
                                                                               "==SQL:" + strSql);
                }

                #endregion
            }
            else if (type == 5)
            {
                #region 根据代理机构编号查询代理机构自增长ID

                try
                {
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        strSql = @"select n_AgencyID  from TCstmr_CoopAgency WHERE s_Code='" + rowName + "'";
                        cmd.CommandText = @"select n_AgencyID  from TCstmr_CoopAgency WHERE s_Code='" + rowName + "'";
                        cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                        using (SqlDataReader reader = cmd.ExecuteReader())
                            while (reader.Read()) sqlColumns.Add(reader[0].ToString());
                    }
                }
                catch (Exception ex)
                {
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("5.GetIDbyName导入错误信息:" + ex.Message +
                                                                               "==SQL:" + strSql);
                }

                #endregion
            }
            else if (type == 6)
            {
                #region 根据代理机构ID查到代理机构地址

                try
                {
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        strSql = @"SELECT TOP 1 n_ID FROM TCstmr_AgencyAddress WHERE n_AgencyID='" + rowName +
                                 "'  ORDER BY n_ID ASC";
                        cmd.CommandText = @"SELECT TOP 1 n_ID FROM TCstmr_AgencyAddress WHERE n_AgencyID='" + rowName +
                                          "'  ORDER BY n_ID ASC";
                        cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                        using (SqlDataReader reader = cmd.ExecuteReader())
                            while (reader.Read()) sqlColumns.Add(reader[0].ToString());
                    }
                }
                catch (Exception ex)
                {
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("6.GetIDbyName导入错误信息:" + ex.Message +
                                                                               "==SQL:" + strSql);
                }

                #endregion
            }
            else if (type == 7)
            {
                #region 根据申请号查找案件ID

                try
                {
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        strSql = @"SELECT n_CaseID FROM dbo.TCase_Base WHERE  s_AppNo='" + rowName +
                                 "'  ORDER BY n_CaseID ASC";
                        cmd.CommandText = @"SELECT n_CaseID FROM dbo.TCase_Base WHERE  s_AppNo='" + rowName +
                                          "'  ORDER BY n_CaseID ASC";
                        cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                        using (SqlDataReader reader = cmd.ExecuteReader())
                            while (reader.Read()) sqlColumns.Add(reader[0].ToString());
                    }
                }
                catch (Exception ex)
                {
                    Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("GetIDbyName导入错误信息:" + ex.Message +
                                                                               "==SQL:" + strSql);
                }

                #endregion
            }
            if (sqlColumns.Count != 0 && !string.IsNullOrEmpty(sqlColumns[0]))
            {
                retName = int.Parse(sqlColumns[0]);
            }
            return retName;
        }

        /// <summary>
        /// 根据传入sql查询固定值
        /// </summary>
        /// <param name="strSql"></param>
        /// <returns></returns>
        private int GetIDbySql(string strSql)
        {
            int retName = 0;
            var sqlColumns = new List<string>();
            strSql = strSql.Replace("\r", "").Replace("\n", "");
            try
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @strSql;
                    cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                    using (SqlDataReader reader = cmd.ExecuteReader())
                        while (reader.Read()) sqlColumns.Add(reader[0].ToString());
                }
            }
            catch (Exception ex)
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("GetIDbySql查询错误信息:" + ex.Message + "==SQL:" +
                                                                           strSql);
            }

            if (sqlColumns.Count != 0 && !string.IsNullOrEmpty(sqlColumns[0]))
            {
                retName = int.Parse(sqlColumns[0]);
            }
            return retName;
        }

        /// <summary>
        /// 根据传入sql查询固定值
        /// </summary>
        /// <param name="strSql"></param>
        /// <returns></returns>
        private object GetTimebySql(string strSql)
        {
            object retName = null;
            var sqlColumns = new List<object>();

            try
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @strSql;
                    cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                    using (SqlDataReader reader = cmd.ExecuteReader())
                        while (reader.Read()) sqlColumns.Add(reader[0].ToString());
                }
            }
            catch (Exception ex)
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("GetTimebySql查询错误信息:" + ex.Message + "==SQL:" +
                                                                           strSql);
            }

            if (sqlColumns.Count != 0)
            {
                retName = sqlColumns[0];
            }
            return retName;
        }

        /// <summary>
        /// 根据传入sql查询固定值
        /// </summary>
        /// <param name="strSql"></param>
        /// <returns></returns>
        private DataTable GetDataTablebySql(string strSql)
        {
            var table = new DataTable();

            try
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @strSql;
                    cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                    using (SqlDataReader reader = cmd.ExecuteReader())
                        table.Load(reader);
                }
            }
            catch (Exception ex)
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("GetDataTablebySql查询错误信息:" + ex.Message +
                                                                           "==SQL:" + strSql);
            }
            return table;
        }

        /// <summary>
        /// 更加SQL执行增、删、改,带有行号
        /// </summary>
        /// <param name="strSql"></param>
        /// <param name="iRow"> </param>
        /// <returns></returns>
        private int InsertbySql(string strSql, int iRow)
        {
            int retName = 0;
            try
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @strSql;
                    cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                    retName = cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("错误信息:" + ex.Message);
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(iRow + "执行SQL:" + strSql);
            }
            return retName;
        }
        /// <summary>
        /// 更加SQL执行增、删、改
        /// </summary>
        private int InsertbySql(string strSql)
        {
            int retName = 0;
            try
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @strSql;
                    cmd.Parameters.Add(new SqlParameter("@name", com_tablename.Text.Trim()));
                    retName = cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("错误信息:" + ex.Message);
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("执行SQL:" + strSql);
            }
            return retName;
        }
        /// <summary>
        /// 法律信息根据传入项目名称返回修改字段
        /// </summary> 
        /// <returns></returns>
        private string GetColum(string name)
        {
            string strSql = "";
            if (name == "PCT公开号")
            {
                strSql = "s_PCTPubNo";
            }
            else if (name == "PCT公开日")
            {
                strSql = "dt_PctPubDate";
            }
            else if (name == "PCT进入日")
            {
                strSql = "dt_PctInNationDate";
            }
            else if (name == "PCT申请号")
            {
                strSql = "s_PCTAppNo";
            }
            else if (name == "PCT申请日")
            {
                strSql = "dt_PctAppDate";
            }
            else if (name == "PCT办登日")
            {
                strSql = "dt_CertfDate";
            }
            else if (name == "公开号")
            {
                strSql = "s_PubNo";
            }
            else if (name == "公开日")
            {
                strSql = "dt_PubDate";
            }
            else if (name == "进入实审日")
            {
                strSql = "dt_SubstantiveExamDate";
            }
            else if (name == "授权公告号")
            {
                strSql = "s_IssuedPubNo";
            }
            else if (name == "授权公告日")
            {
                strSql = "dt_IssuedPubDate";
            }
            return strSql;
        }

        #endregion

        #endregion

        #region 窗体调用方法

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

        /// <summary>
        /// 读取Excel
        /// </summary>
        /// <param name="filename">文件名(含格式)</param>
        /// <param name="index">页数(默认0:第一页)</param>
        private void ReadExcel(string filename, int index)
        {
            //string strConn = "Provider=Microsoft.Jet.OleDb.4.0;" + "data source=" + excelFile + ";Extended Properties='Excel 8.0;'"; //此連接只能操作Excel2007之前(.xls)文件
            string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + filename +
                             ";Extended Properties='Excel 12.0;'"; //此連接可以操作.xls與.xlsx文件
            using (OleDbConnection conn = new OleDbConnection(strConn))
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
                    using (OleDbDataAdapter da = new OleDbDataAdapter(cmd))
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

        #endregion
    }
}