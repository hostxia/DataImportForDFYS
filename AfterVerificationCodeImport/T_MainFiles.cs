using System;
using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport
{
    class T_MainFiles
    {
        readonly DBHelper _dbHelper = new DBHelper();

        public void OADataInsertTMainFile(int nCaseID, int rowID, string Name, string TableName, string dtSendDate, int ObjectType, string sAbstact, string sClientGov, string sIOType, string commDB, SqlConnection _connection)
        {
            
            string strSql =
                "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,ObjectType,dt_EditDate,s_Name,s_Abstact,dt_CreateDate,s_ClientGov,s_IOType,dt_SendDate)" +
                "VALUES  ('" + TableName +"-"+ rowID + "'," + ObjectType + ",'" + DateTime.Now.AddMonths(-1) + "','" + Name.Replace("'", "''") + "','" + sAbstact.Replace("'", "''") + "','" + DateTime.Now.AddMonths(-1) + "','" + sClientGov + "','" + sIOType + "','" + dtSendDate + "')";

            string strSqlS =
                "SELECT top 1 n_FileID FROM dbo.T_MainFiles WHERE  s_IOType='" + sIOType + "' AND s_ClientGov='" +
                sClientGov +
                "' and s_Name='" + Name.Replace("'", "''") + "' and s_Abstact='" +
                sAbstact.Replace("'", "''") + "' and dt_SendDate='" + dtSendDate +
                "' and ObjectType=" + ObjectType + " order by n_FileID desc ";
            int nFileID = 0;

            if (_dbHelper.InsertbySql(strSql, rowID, commDB, _connection) > 0)
            {
                nFileID = _dbHelper.GetbySql(strSqlS, commDB, _connection); 
            }

            if (nFileID > 0)
            {
                int fCount = 0;
                string sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_FileInCase WHERE n_FileID=" + nFileID +
                             " and n_CaseID=" + nCaseID;
                int sumFileInCase = _dbHelper.GetbySql(sql, commDB, _connection);
                if (sumFileInCase <= 0)
                {
                    sql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                          "VALUES  (" + nCaseID + " ," + nFileID + ",'Y')";
                    int fileInCase = _dbHelper.InsertbySql(sql, rowID, commDB, _connection); //记录文件、案件与程序的关系表 
                    if (fileInCase <= 0)
                    {
                        _dbHelper.InsertLog(nCaseID, "", rowID, TableName, TableName + rowID, "T_FileInCase插入数据错误", sql.Replace("'", "''"), commDB, _connection);
                    }
                    fCount = fileInCase;
                }
                else
                {
                    fCount = sumFileInCase; 
                }
                int nGovOfficeID = 0;
                if (sClientGov.Equals("O")) //s_ClientGov C: 客户 O: 官方
                {
                    nGovOfficeID = 21; //中国国家知识产权局
                }
                if (fCount > 0)
                {
                    if (!Name.Equals("官方来文"))
                    {
                        string sqlS =
                            "INSERT INTO dbo.T_OutFiles( n_FileID ,n_CheckedOutBy , n_GovOfficeID , s_FileStatus, dt_StatusDate ,dt_WriteDate ," +
                            "n_WriterID , n_SubmiterID ,  n_PrintNum , n_PageNum ,n_ReFileID  ,n_Count ,s_FileType ,n_LatestCheckInfoID)" +
                            "VALUES  (" + nFileID + ",0 ," + nGovOfficeID + " ,'W' ,'" + DateTime.Now + "' ,'" +
                            DateTime.Now + "',0 ,0 ,1,0 ,0,0 ,'new',0 )";
                        sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_OutFiles WHERE n_FileID=" + nFileID;
                        int sumcount = _dbHelper.GetbySql(sql, commDB, _connection);
                        if (sumcount <= 0)
                        {
                            int outFiles = _dbHelper.InsertbySql(sqlS, rowID, commDB, _connection);
                            if (outFiles <= 0)
                            {
                                _dbHelper.InsertLog(nCaseID, "", rowID, TableName, TableName + rowID, "T_OutFiles插入数据错误", sqlS.Replace("'", "''"), commDB, _connection);
                            }
                        }
                    }
                    else
                    {
                        string sqlS =
                          "INSERT INTO dbo.T_InFiles( n_FileID,n_FileCodeID,n_GovOfficeID,s_OFileStatus,s_Distribute,s_Note)" +
                           "VALUES  (" + nFileID + ",0," + nGovOfficeID + " ,'N','Y','OA数据补充导入-OA收到日')";
                        sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_InFiles WHERE n_FileID=" + nFileID;
                        int sumcount = _dbHelper.GetbySql(sql, commDB, _connection);
                        if (sumcount <= 0)
                        {
                            int outFiles = _dbHelper.InsertbySql(sqlS, rowID, commDB, _connection);
                            if (outFiles <= 0)
                            {
                                _dbHelper.InsertLog(nCaseID, "", rowID, TableName, TableName + rowID, "T_InFiles插入数据错误", sqlS.Replace("'", "''"), commDB, _connection);
                            }
                        } 
                    }
                }
            }
            else
            {
                _dbHelper.InsertLog(nCaseID, "", rowID, TableName, TableName + rowID, "未查询到相关的文件往来", strSqlS.Replace("'", "''") + "   \r\n" + strSql.Replace("'", "''"), commDB, _connection);
            }
        }

        //案件与文件关系
        public void InserIntoTFileInCase(int hkNum, int nFileID, int rowid, string sNo, string tabName, string commDB, SqlConnection _connection)
        {
            string strSql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                            "VALUES  (" + hkNum + " ," + nFileID + ",'Y')";
            if (_dbHelper.InsertbySql(strSql, rowid, commDB, _connection) <= 0)
            {
                _dbHelper.InsertLog(hkNum, sNo, rowid, tabName, tabName + "-" + rowid, "增加数据失败[InserIntoTFileInCase]", strSql.Replace("'", "''"), commDB, _connection);
            }
        }

        //文件信息基础表 官方 发文
        public int InserIntoTMainFiles(string idsName, DateTime insertTime, string time, int rowid, int hkNum, string sNo, int ObjectTypeID, string tabName, string commDB, SqlConnection _connection)
        {
            string strSql =
                "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,s_Abstact,ObjectType,s_SendMethod,dt_EditDate,s_Name,dt_SendDate,dt_CreateDate,s_IOType,s_ClientGov)" +
                "VALUES  ('" + tabName + rowid + "',''," + ObjectTypeID + ",'其他','" + insertTime + "','" + idsName + "','" + time + "','" + insertTime + "','O','O') ";
            int Num = _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
            if (Num <= 0)
            {
                _dbHelper.InsertLog(hkNum, sNo, rowid, tabName, tabName + "-" + rowid, "增加数据失败[InserIntoTMainFiles]", strSql.Replace("'", "''"), commDB, _connection);
            }
            return Num;
        }

        //案件与发文关系 官方
        public int InserIntoTOutFiles(int nGovOfficeID, int nFileID, int rowid, string commDB, SqlConnection _connection)
        {
            string strSql =
                "INSERT INTO dbo.T_OutFiles( n_FileID ,n_CheckedOutBy , n_GovOfficeID , s_FileStatus, dt_StatusDate ,dt_WriteDate ," +
                "n_WriterID , n_SubmiterID ,  n_PrintNum , n_PageNum ,n_ReFileID  ,n_Count ,s_FileType ,n_LatestCheckInfoID)" +
                "VALUES  (" + nFileID + ",0 ," + nGovOfficeID + " ,'W' ,'" + DateTime.Now + "' ,'" + DateTime.Now +
                "',0 ,0 ,1,0 ,0,0 ,'new',0 )";
             
            string sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_OutFiles WHERE n_FileID=" +
                         nFileID;
            if (_dbHelper.GetbySql(sql, commDB, _connection) <= 0)
            {
                return _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
            }
            return 0;
        }

        //案件与案件关系
        public void InsertTCaseCaseRelative(int caseIDA, int caseIDB, int n_ID, int s_MasterSlaveRelation, int rowid, string commDB, SqlConnection _connection)
        {
            //TCase_CaseRelative-案件关系表，其中n_CaseIDA与n_CaseIDB为两个案件，如果s_MasterSlaveRelation为1时，表示A为主案件；0表示B为主案件；-1为两者为平级，不分主从。n_CodeRelativeID为对应的关联关系配置ID
            string strSql =
                " INSERT INTO  dbo.TCase_CaseRelative ( n_CaseIDA ,  n_CaseIDB , dt_CreateDate , dt_EditDate , s_MasterSlaveRelation , n_CodeRelativeID )" +
                " VALUES  ( " + caseIDA + " , " + caseIDB + " ,  GETDATE() , GETDATE() ,  " + s_MasterSlaveRelation +
                " ,  " + n_ID + ")";
            _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
        }

        //发文主表
        public void InTMainFiles(string Name, string dtSendDate, int numHk, int rowid, int agent, int caseNo, string sClientGov, int OutFileID, string commDB, SqlConnection _connection)
        {
            DateTime insertTime = DateTime.Now;
            string strSql =
                       "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,ObjectType, dt_EditDate,s_IOType, s_ClientGov,s_SendMethod ,s_Name ,s_Abstact,dt_CreateDate,dt_SendDate";
            string strSql2 = "VALUES  ('香港-专利数据补充导入'," + OutFileID + ",'" + insertTime + "','O','" + sClientGov + "','其他','" + Name + "','" + Name + "','" + insertTime + "'," + dtSendDate;
            if (caseNo > 0)
            {
                strSql += ",n_ClientID";
                strSql2 += "," + caseNo;
            }
            strSql += ")";
            strSql2 += ")";
            string strSqlS =
                       "SELECT n_FileID FROM dbo.T_MainFiles WHERE   s_IOType='O' and  s_ClientGov='" + sClientGov + "' and s_SendMethod='其他' and s_Name='" + Name + "'  and s_Abstact='" +
                       Name + "' and ObjectType=" + OutFileID + " and dt_SendDate=" + dtSendDate;
            if (caseNo > 0)
            {
                strSqlS += " and n_ClientID=" + caseNo;
            }
            int nFileID2 = _dbHelper.GetbySql(strSqlS, commDB, _connection);
            if (nFileID2 > 0)
            {
                InTFileInCase(numHk, nFileID2, 0, agent, commDB, _connection);
            }
            else
            {
                _dbHelper.InsertbySql(strSql + strSql2, rowid, commDB, _connection);
                nFileID2 = _dbHelper.GetbySql(strSqlS, commDB, _connection);
                if (nFileID2 > 0)
                {
                    InTFileInCase(numHk, nFileID2, 0, agent, commDB, _connection);
                }
            }
        }
        //案件和发文表关系
        private void InTFileInCase(int numHk, int nFileID, int i, int agent, string commDB, SqlConnection _connection)
        {
            int fCount = 0;
            string sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_FileInCase WHERE n_FileID=" + nFileID +
                         " and n_CaseID=" + numHk;
            int sumFileInCase = _dbHelper.GetbySql(sql, commDB, _connection);
            if (sumFileInCase <= 0)
            {
                sql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                      "VALUES  (" + numHk + " ," + nFileID + ",'Y')";
                int fileInCase = _dbHelper.InsertbySql(sql, i, commDB, _connection); //记录文件、案件与程序的关系表 
                if (fileInCase <= 0)
                {
                    _dbHelper.InsertLog(numHk, "", i, "香港-专利数据补充导入", "香港-专利数据补充导入-" + i, "T_FileInCase插入数据错误", sql.Replace("'", "''"), commDB, _connection);
                }
                fCount = fileInCase;
            }
            else
            {
                fCount = sumFileInCase;
            }
            if (fCount > 0)
            {
                InTOutFiles(nFileID, i, agent, numHk, commDB, _connection);
            }
        }
        //发文表
        private void InTOutFiles(int nFileID, int i, int agent, int numHk, string commDB, SqlConnection _connection)
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
            int sumcount = _dbHelper.GetbySql(sql, commDB, _connection);
            if (sumcount <= 0)
            {
                int outFiles = _dbHelper.InsertbySql(sqlS + strSql2, i, commDB, _connection);
                if (outFiles <= 0)
                {
                    _dbHelper.InsertLog(numHk, "", i, "香港-专利数据补充导入", "香港-专利数据补充导入-" + i, "T_FileInCase插入数据错误", sqlS + strSql2.Replace("'", "''"), commDB, _connection);
                }
            }
        }


        #region 国内-收文
        public int InsertFileIn(int rowid, DataRow dr, int InFileID, string commDB, SqlConnection _connection)
        {
            int result = 0;
            string sNo = dr["我方卷号"].ToString().Trim();
            int numHk = _dbHelper.GetIDbyName(sNo, 2, _connection);
            if (numHk.Equals(0))
            {
                _dbHelper.InsertLog(0, sNo, rowid, "国内-收文数据 ", "国内-收文数据-" + rowid, "未找到“我方卷号”为：" + sNo, "", commDB, _connection);
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
                    string content = "客户来文   发件人：" + dr["发件人"].ToString().Replace("\r", "").Replace("\n", "");
                    remark = content;
                }
                DateTime insertTime = DateTime.Now.AddMonths(-1);
                try
                {
                    int nClientID = 21;
                    var strSql = "";
                    if (sClientGov.Equals("C"))
                    {
                        nClientID =
                            _dbHelper.GetbySql("  SELECT top 1 n_ClientID FROM TCstmr_Client WHERE s_Name='" +
                                       dr["发件人"].ToString().Trim().Replace("'", "''") + "'", commDB, _connection);
                        if (nClientID == 0)
                        {
                            strSql =
                                 "INSERT INTO dbo.TCstmr_Client(s_Name) " +
                                 "VALUES('" + dr["发件人"].ToString().Trim().Replace("'", "''").Replace("\r", "").Replace("\n", "") + "')";
                            _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                            nClientID =
                            _dbHelper.GetbySql("  SELECT top 1 n_ClientID FROM TCstmr_Client WHERE s_Name='" +
                                      dr["发件人"].ToString().Trim().Replace("'", "''") + "'", commDB, _connection);
                        }
                    }
                    string con = dr["内容"].ToString().Replace("'", "''").Trim() == ""
                                     ? "无名称"
                                     : dr["内容"].ToString().Replace("'", "''");

                    strSql =
                       "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,ObjectType,s_Status,dt_EditDate,s_IOType, s_ClientGov,s_SendMethod ,s_Name ,s_Abstact,dt_CreateDate ";
                    var stesql2 = "VALUES  ('国内-收文数据导入" + rowid + "'," + InFileID + ",'Y','" + insertTime + "','I','" + sClientGov + "','" +
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
                     
                    string strSqlS =
                        "SELECT top 1 n_FileID FROM dbo.T_MainFiles WHERE  s_IOType='I' and  s_ClientGov='" +
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
                    int nFileID = 0; 
                    using (SqlCommand cmd = _connection.CreateCommand())
                    {
                        cmd.CommandText = strSql + stesql2; 
                        if (cmd.ExecuteNonQuery() > 0)
                        {
                            nFileID = _dbHelper.GetbySql(strSqlS, commDB, _connection); 
                        }
                    } 
                    int nGovOfficeID = 0;
                    if (sClientGov.Equals("O")) //s_ClientGov C: 客户 O: 官方
                    {
                        nGovOfficeID = 21; //中国国家知识产权局
                    }
                    if (nFileID > 0)
                    {
                        string sql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                                     "VALUES  (" + numHk + " ," + nFileID + ",'Y')";
                        strSql = "SELECT COUNT(*) AS sumNum FROM dbo.T_FileInCase WHERE n_CaseID=" + numHk +
                                 " AND n_FileID=" + nFileID;
                        int sumNumFileInCase = _dbHelper.GetbySql(strSql, commDB, _connection);
                        if (sumNumFileInCase <= 0)
                        {
                            _dbHelper.InsertbySql(sql, rowid, commDB, _connection);
                        }
                        else
                        {
                           Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer(rowid + "插入T_InFiles表数已存在" + strSql);
                        }
                        string sGetCertificater = string.Empty;
                        if (dr["处理人"] != null && !string.IsNullOrEmpty(dr["处理人"].ToString()))
                        {
                            sGetCertificater = " and dt_ReceiveDate='" + dr["处理人"] + "'";
                        }

                        strSql = "SELECT COUNT(*) AS sumNum FROM dbo.T_InFiles WHERE n_FileID=" + nFileID +
                                 " and n_GovOfficeID=" + nGovOfficeID + " and s_Distribute='Y'  and s_Note='N' and s_GetCertificater='"+sGetCertificater+"'";
                        int sumNumInFiles = _dbHelper.GetbySql(strSql, commDB, _connection);
                        if (sumNumInFiles <= 0)
                        {
                            sql =
                                "INSERT INTO dbo.T_InFiles(n_FileID,n_GovOfficeID,dt_TransmitDate,dt_GetCertificatedate,s_Distribute,s_Note，s_GetCertificater)" +
                                "VALUES  (" + nFileID + " ," + nGovOfficeID + ",'" + DateTime.Now + "','" +
                                DateTime.Now + "','Y','N'，'" + sGetCertificater + "')";
                            _dbHelper.InsertbySql(sql, rowid, commDB, _connection);
                        }
                    }
                    else
                    {
                        _dbHelper.InsertLog(0, sNo, rowid, "国内-收文数据 ", "国内-收文数据-" + rowid, "插入T_MainFiles表数失败", strSql + stesql2, commDB, _connection); 
                    }
                    result = 1;
                }
                catch (Exception ex)
                {
                    _dbHelper.InsertLog(0, sNo, rowid, "国内-收文数据", "国内-收文数据-" + rowid, "插入T_MainFiles表数据错误，“我方卷号”为:" + dr["我方卷号"].ToString().Trim() + "  错误信息:" + ex.Message, "", commDB, _connection); 
                }
            }
            return result;
        }

        #endregion 
    }
}
