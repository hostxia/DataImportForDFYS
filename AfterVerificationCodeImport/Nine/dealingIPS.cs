using System;
using System.Data;
using System.Data.SqlClient;
using AfterVerificationCodeImport.Four;

namespace AfterVerificationCodeImport.Nine
{
    class dealingIPS
    {
        readonly DBHelper _dbHelper = new DBHelper();

        #region 要求/发文
        public int sGetType(DataRow dataRow, int row, int OutFileID, string commDB, SqlConnection _connection)
        {
            if (dataRow["栏目名称"] != null && dataRow["我方文号"] != null)
            {
                if (dataRow["导入形式"] != null &&
                    (dataRow["导入形式"].ToString().Equals("发客户文") || dataRow["导入形式"].ToString().Equals("发官方文")))
                {
                    return InsertOutFile(dataRow, row, OutFileID, commDB, _connection);
                }
                else if (dataRow["导入形式"] != null && dataRow["导入形式"].ToString().Equals("案件客户要求"))
                {
                    return InsertOnlyCaseDemand(dataRow, row, commDB, _connection);
                }
            }
            else
            {
                _dbHelper.InsertLog(0, "", row, "IPS栏目名称", "IPS栏目名称-" + row, "栏目名称或者我方文号为空", "", commDB, _connection);
            }
            return 0;
        }

        private int InsertOutFile(DataRow dataRow, int row, int OutFileID, string commDB, SqlConnection _connection)
        {
            if (dataRow["我方卷号"] != null && dataRow["导入形式"] != null)
            {
                string sNo = dataRow["我方卷号"].ToString().Trim();
                int numHk = _dbHelper.GetIDbyName(sNo, 2,_connection);
                if (numHk <= 0)
                {
                    _dbHelper.InsertLog(0, "", row, "IPS栏目名称", "IPS栏目名称-" + row, "导入形式或者我方文号为空", "", commDB, _connection);
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
                    if (_dbHelper.InsertbySql(strSql, row, commDB, _connection) > 0)
                    {
                        nFileID = _dbHelper.GetbySql(sqlS, commDB, _connection);
                    }
                    if (nFileID > 0)
                    {
                        string sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_FileInCase WHERE n_FileID=" + nFileID +
                                     " and n_CaseID=" + numHk;
                        int sumFileInCase = _dbHelper.GetbySql(sql, commDB, _connection);
                        if (sumFileInCase <= 0)
                        {
                            sql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                                  "VALUES  (" + numHk + " ," + nFileID + ",'Y')";
                            sumFileInCase = _dbHelper.InsertbySql(sql, row, commDB, _connection); //记录文件、案件与程序的关系表 
                            if (sumFileInCase <= 0)
                            {
                                _dbHelper.InsertLog(numHk, sNo, row, "发文官方/发文客户/案件个案要求-", "发文官方/发文客户/案件个案要求--" + row, "T_FileInCase插入数据错误：" + sNo, sql.Replace("'", "''"), commDB, _connection); 
                            }
                        }
                        else
                        {
                            _dbHelper.InsertLog(numHk, sNo, row, "发文官方/发文客户/案件个案要求-", "发文官方/发文客户/案件个案要求--" + row, "已存在案件和文件关系，无需插入：" + sNo, sql.Replace("'", "''"), commDB, _connection); 
                        }
                        int nGovOfficeID = 0;
                        if (sClientGov.Equals("O")) //s_ClientGov C: 客户 O: 官方
                        {
                            nGovOfficeID = 21; //中国国家知识产权局
                        }
                        if (sumFileInCase > 0)
                        {
                            sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_OutFiles WHERE n_FileID=" + nFileID;
                            int sumcount = _dbHelper.GetbySql(sql, commDB, _connection);
                            if (sumcount <= 0)
                            {
                                sqlS =
                                    "INSERT INTO dbo.T_OutFiles( n_FileID ,n_CheckedOutBy , n_GovOfficeID , s_FileStatus, dt_StatusDate ,dt_WriteDate ," +
                                    "n_WriterID , n_SubmiterID ,  n_PrintNum , n_PageNum ,n_ReFileID  ,n_Count ,s_FileType ,n_LatestCheckInfoID)" +
                                    "VALUES  (" + nFileID + ",0 ," + nGovOfficeID + " ,'W' ,'" + DateTime.Now + "' ,'" +
                                    DateTime.Now + "',0 ,0 ,1,0 ,0,0 ,'new',0 )";

                                int outFiles = _dbHelper.InsertbySql(sqlS, row, commDB, _connection);
                                if (outFiles <= 0)
                                {
                                    _dbHelper.InsertLog(numHk, sNo, row, "发文官方/发文客户/案件个案要求-", "发文官方/发文客户/案件个案要求--" + row, "T_OutFiles插入数据错误：" + sNo, sqlS.Replace("'", "''"), commDB, _connection); 
                                }
                            }
                        }
                    }
                    return 1;
                }
            }
            return 0;
        }

        private int InsertOnlyCaseDemand(DataRow dataRow, int row, string commDB, SqlConnection _connection)
        {
            string sNo = dataRow["我方卷号"].ToString().Trim();
            int numHk = _dbHelper.GetIDbyName(sNo, 2,_connection);
            string content = dataRow["栏目名称"].ToString().Replace("'", "''") + "：" + dataRow["栏目内容"].ToString().Replace("'", "''");
            string Sql = "SELECT n_ID  FROM T_Demand where  s_Title='" + content + "' and n_CaseID=" + numHk + " and s_Description='" + dataRow["备注"].ToString().Replace("'", "''") + "'";
            int n_ID = _dbHelper.GetbySql(Sql, commDB, _connection);
            if (n_ID <= 0 && numHk > 0)
            {
                Sql = "INSERT INTO dbo.T_Demand(s_sourcetype1,dt_EditDate,s_Title,dt_CreateDate,n_CaseID,s_Description)" +
                      "VALUES  ('案件要求','" + DateTime.Now + "','" + content + "','" + DateTime.Now + "'," + numHk + ",'" + dataRow["备注"].ToString().Replace("'", "''") + "')";
                int ResultNum = _dbHelper.InsertbySql(Sql, row, commDB, _connection);
                if (ResultNum.Equals(0))
                {
                    _dbHelper.InsertLog(0, "", row, "IPS栏目名称", "IPS栏目名称-" + row, "发文记录和案件要求-增加案件要求失败[InsertOnlyCaseDemand]", Sql.Replace("'", "''"), commDB, _connection);
                }
            }
            else
            {
                _dbHelper.InsertLog(0, "", row, "IPS栏目名称", "IPS栏目名称-" + row, "未找到我方卷号或者已存在案件要求", "", commDB, _connection);
            }
            return 0;
        }

        #endregion 

        #region 档案位置
        public int InsertFileLocation(DataRow dataRow, int row, string commDB, SqlConnection _connection)
        {
            string sNo = dataRow["我方文号"].ToString().Trim();
            int numHk = _dbHelper.GetIDbyName(sNo, 2,_connection);
            if (numHk > 0)
            {
                //当前案件是否是借出状态
                string strSql = "SELECT s_ArchiveStatus FROM dbo.TCase_Base WHERE n_CaseID=" + numHk;
                string sArchiveStatus = _dbHelper.GetStringbySql(strSql, _connection);

                if (sArchiveStatus.ToUpper().Equals("I"))
                {
                    strSql = "SELECT n_ID FROM  dbo.TCode_Employee  WHERE s_Name='" + dataRow["借卷人"].ToString().Trim().Replace("'", "''") + "' OR s_InternalCode='" + dataRow["借卷人"].ToString().Trim().Replace("'", "''") + "'";
                    int nUserId = _dbHelper.GetbySql(strSql, commDB, _connection);
                    if (nUserId > 0)
                    {
                        strSql = "UPDATE dbo.TCase_Base SET s_ArchiveStatus='O',s_ArchivePosition=" + nUserId + " WHERE n_CaseID=" + numHk;
                        strSql += "  INSERT INTO dbo.TCase_ArchivesHistory(n_CaseID ,n_BorrowerID ,dt_BorrowerTime,s_Notes ,dt_CreateTime)" +
                                  "  VALUES  (" + numHk + "," + nUserId + ",'" + dataRow["借卷时间"].ToString() + "','借出：','" + DateTime.Now + "')";
                        return _dbHelper.InsertbySql(strSql, row, commDB, _connection);
                    }
                    else
                    {
                        _dbHelper.InsertLog(numHk, sNo, row, "档案位置", "档案位置-" + row, "此案件借卷人未查到", "", commDB, _connection);
                    }
                }
                else
                {
                    _dbHelper.InsertLog(numHk, sNo, row, "档案位置", "档案位置-" + row, "此案件未借出状态，无法再次外借", "", commDB, _connection);
                }
            }
            else
            {
                _dbHelper.InsertLog(0, sNo, row, "档案位置", "档案位置-" + row, "未找到“我方卷号”为：" + sNo, "", commDB, _connection);
            }
            return 0;
        }
        #endregion

        #region 澳门案件
        public int InsertMacaoApplication(DataRow dr, int rowid, string commDB, SqlConnection _connection)
        { 
            string sNo = dr["EARLIER"].ToString().Trim();
            int HKNum = _dbHelper.GetIDbyName(sNo, 2, _connection);
            int MacaoID = _dbHelper.GetIDbyName(dr["LATER"].ToString().Trim(), 2, _connection);
            if (HKNum.Equals(0))
            {
                _dbHelper.InsertLog(0, sNo, rowid, "澳门案件", "澳门案件-" + rowid, "未找到“我方卷号”为：" + sNo, "", commDB, _connection);
                return 0;
            }
            else
            {
                const string strSql = "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='澳门延伸' AND s_MasterName='国内母案' AND s_SlaveName='澳门延伸' AND s_IPType='P'";
                int n_ID = _dbHelper.GetbySql(strSql, commDB, _connection);
                var _dealingCaseToCase = new dealingCaseToCase();
                _dealingCaseToCase.InsertIntoLaw(HKNum, MacaoID, n_ID, rowid, "", commDB, _connection);
                UpdateHongKangandMacao(HKNum, MacaoID, rowid, commDB, _connection);//同步香港和澳门案件同族
                return 1;
            }
        }
        private void UpdateHongKangandMacao(int ncaseIDF, int nCaseIDC, int rowid, string commDB, SqlConnection _connection)
        { 
            var strSql = "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='香港申请' AND s_MasterName='国内母案' AND s_SlaveName='香港案' AND s_IPType='P'";
            //案件关系配置表
            int nIDHongKang = _dbHelper.GetbySql(strSql, commDB, _connection);

            strSql = "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='同族专利' AND s_MasterName='同族' AND s_SlaveName='同族' AND s_IPType='P'";
            int nIDTong = _dbHelper.GetbySql(strSql, commDB, _connection);

            strSql = "SELECT n_CaseIDA FROM TCase_CaseRelative WHERE n_CaseIDB=" + ncaseIDF + " AND n_CodeRelativeID=" + nIDHongKang;//查找所有香港申请案件
            DataTable table = _dbHelper.GetDataTablebySql(strSql, _connection);
            for (int i = 0; i < table.Rows.Count; i++)
            {
                var _dealingCaseToCase = new dealingCaseToCase();
                _dealingCaseToCase.InsertIntoLaw(nCaseIDC, int.Parse(table.Rows[i]["n_CaseIDA"].ToString()), nIDTong, rowid, "同族", commDB, _connection); 
            }
        }
        #endregion 
    }
}
