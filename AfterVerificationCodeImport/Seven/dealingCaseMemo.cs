using System;
using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport.Seven
{
    class dealingCaseMemo
    {
        readonly DBHelper _dbHelper = new DBHelper();

        public int Case_Memo(int rowid, DataRow dr, int OutFileID, string commDB, SqlConnection _connection)
        {
            int result = 0;
            string sNo = dr["我方卷号"].ToString().Trim();
            int hkNum = _dbHelper.GetIDbyName(sNo, 2,_connection);
            if (hkNum.Equals(0))
            {
                _dbHelper.InsertLog(0, sNo, rowid, "国外-国外库时限备注表 ", "国外-国外库时限备注表 -" + rowid, "未找到“我方卷号”为：" + sNo, "", commDB, _connection);
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

                    int nFileID = _dbHelper.GetbySql(strSql1, commDB, _connection);
                    if (nFileID > 0)
                    {
                        var _tMainFiles = new T_MainFiles();
                        string strSql2 = " SELECT n_FileID FROM dbo.T_OutFiles WHERE n_FileID=" + nFileID;
                        int nFileIDOut = _dbHelper.GetbySql(strSql2, commDB, _connection);
                        if (nFileIDOut > 0) //发件表和文件主表关联
                        {
                            string strSql3 = "  select n_ID FROM T_FileInCase  WHERE n_FileID =" + nFileID +
                                             " AND n_CaseID=" + hkNum;
                            int nID = _dbHelper.GetbySql(strSql3, commDB, _connection);
                            if (nID <= 0) //发件和无案件关联
                            {
                                _tMainFiles.InserIntoTFileInCase(hkNum, nFileID, rowid, sNo, "国外-国外库时限备注表", commDB, _connection);
                            }
                        }
                        else
                        {
                            if (_tMainFiles.InserIntoTOutFiles(0, nFileID, rowid, commDB, _connection) > 0)
                            {
                                _tMainFiles.InserIntoTFileInCase(hkNum, nFileID, rowid, sNo, "国外-国外库时限备注表", commDB, _connection);
                            }
                        }
                    }
                    else
                    {

                        var insertTime = DateTime.Now.AddMonths(-1);
                        string strSql =
                            "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,ObjectType,s_Status,s_SendMethod,dt_EditDate,s_Name,dt_CreateDate,s_IOType,s_ClientGov,s_Abstact ";
                        string strsql2 = "VALUES  ('国外-国外库时限备注表" + rowid + "'," + OutFileID + ",'Y','其他','" + insertTime + "','监视类型：" +
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
                        _dbHelper.InsertbySql(strSql + strsql2, rowid, commDB, _connection);
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
                        nFileID = _dbHelper.GetbySql(strSql + " order by n_FileID desc ", commDB, _connection);
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
                            if (_dbHelper.GetbySql(sql, commDB, _connection) <= 0)
                            {
                                string AddSql = strSql + valueSql;
                                _dbHelper.InsertbySql(AddSql, rowid, commDB, _connection);

                            }
                            strSql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                                     "VALUES  (" + hkNum + " ," + nFileID + ",'Y')";

                            string sqlInCase = "SELECT COUNT(*) AS sumcount FROM dbo.T_FileInCase WHERE n_CaseID=" + hkNum + " and n_FileID=" + nFileID;
                            if (_dbHelper.GetbySql(sqlInCase, commDB, _connection) <= 0)
                            {
                                _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                            }
                        }
                    }
                }
                else
                {
                    _dbHelper.InsertLog(0, sNo, rowid, "国外-国外库时限备注表 ", ".国外-国外库时限备注表 -" + rowid, "完成日为空，无需处理", "", commDB, _connection);
                }
            }
            return result;
        }


        #region 备注
        public int InsertTCaseMemo(DataRow dataRow, int row, string commDB, SqlConnection _connection)
        {
            string sNo = dataRow["我方卷号"].ToString().Trim();
            int numHk = _dbHelper.GetIDbyName(sNo, 2,_connection);
            string Sql = "SELECT n_ID  FROM TCase_Memo where  s_Memo='" + dataRow["备注内容"].ToString().Replace("'", "''") + "' and n_CaseID=" + numHk + " and s_Type='" + dataRow["备注类型"].ToString().Replace("'", "''") + "'";
            int n_ID = _dbHelper.GetbySql(Sql.Replace("'", "''"), commDB, _connection);
            if (numHk > 0)
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
                        int ResultNum = _dbHelper.InsertbySql(Sql.Replace("'", "''"), row, commDB, _connection);
                        if (ResultNum.Equals(0))
                        {
                            _dbHelper.InsertLog(numHk, sNo, row, "备注", "备注-" + row, "增加备注失败[InsertTCaseMemo]" + sNo, Sql.Replace("'", "''"), commDB, _connection);
                        }
                    }
                }
            }
            else
            {
                _dbHelper.InsertLog(0, sNo, row, "备注", "备注-" + row, "未找到“我方卷号”为：" + sNo, "", commDB, _connection);
            }
            return 0;
        }
        #endregion
    }
}
