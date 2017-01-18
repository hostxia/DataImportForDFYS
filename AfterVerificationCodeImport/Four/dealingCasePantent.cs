using System;
using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport.Four
{
    class dealingCasePantent
    {
        readonly DBHelper _dbHelper = new DBHelper();

        #region 国内-专利数据补充导入
        public int InsertPatented(int _rowid, DataRow dr, string commDB, SqlConnection _connection)
        {
            string Sql = "";
            int Result = 0;
            string s_No = dr["我方卷号"].ToString().Trim();
            int HkNum = _dbHelper.GetIDbyName(s_No, 2,_connection);
            if (HkNum.Equals(0))
            {
                _dbHelper.InsertLog(0, s_No, _rowid, "国内-专利数据补充导入", "国内-专利数据补充导入-" + _rowid, "未找到“我方卷号”为：" + s_No, "", commDB, _connection);
                return Result;
            }
            if (HkNum != 0)
            {

                #region 更新TPCase_Patent
                try
                {
                    if (dr["分案申请提交日"] != null && !string.IsNullOrEmpty(dr["分案申请提交日"].ToString()))
                    {
                        Sql += "UPDATE TPCase_Patent set dt_DivSubmitDate='" + dr["分案申请提交日"] + "'";
                    }
                    if (dr["提实审日期"] != null && !string.IsNullOrEmpty(dr["提实审日期"].ToString()))
                    {
                        if (Sql.Contains("UPDATE TPCase_Patent"))
                        {
                            Sql += ",dt_RequestSubmitDate='" + dr["提实审日期"] + "'";
                        }
                        else
                        {
                            Sql += "UPDATE TPCase_Patent set dt_RequestSubmitDate='" + dr["提实审日期"] + "'";
                        }
                    }
                    if (!string.IsNullOrEmpty(Sql))
                    {
                        Sql += " WHERE n_CaseID=" + HkNum;
                        using (SqlCommand cmd = _connection.CreateCommand())
                        {
                            cmd.CommandText = Sql;
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
                catch (Exception ex)
                {
                    _dbHelper.InsertLog(HkNum, s_No, _rowid, "国内-专利数据补充导入", "国内-专利数据补充导入-" + _rowid, "更新TPCase_Patent信息错误：" + ex.Message.Replace("'", "''"), Sql.Replace("'", "''"), commDB, _connection);
                }
                #endregion 

                #region 更新TPCase_LawInfo
                try
                {
                    Sql = "UPDATE TPCase_LawInfo  set  s_PCTPubNo='" + dr["PCT公开号"] +
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
                    using (SqlCommand cmd = _connection.CreateCommand())
                    {
                        cmd.CommandText = Sql; 
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    _dbHelper.InsertLog(HkNum, s_No, _rowid, "国内-专利数据补充导入", "国内-专利数据补充导入-" + _rowid, "更新TPCase_LawInfo信息错误：" + ex.Message, Sql.Replace("'", "''"), commDB, _connection);
                }
                #endregion

                #region  增加要求
                if (dr["年费备注"] != null && !string.IsNullOrEmpty(dr["年费备注"].ToString()))
                {
                    InsertPatentSemand(dr["年费备注"].ToString().Replace("'", "''").Replace("\r", "").Replace("\n", ""), HkNum, _rowid, commDB, _connection);
                }
                if (dr["客户指示提实审"] != null && !string.IsNullOrEmpty(dr["客户指示提实审"].ToString()) && dr["客户指示提实审"].ToString().Trim().ToUpper().Equals("Y"))
                {
                    //添加GA01 要求
                    var sql = "select n_ID from TCode_Demand where s_sysdemand='GA01'";
                    var table = _dbHelper.GetDataTablebySql(sql,_connection);
                    if (table.Rows.Count > 0)
                    {
                        sql = "select n_ID from T_Demand where n_CodeDemandID=" + table.Rows[0]["n_ID"].ToString();
                        var num = _dbHelper.GetbySql(sql, commDB, _connection);
                        InsertCaseE(num, _rowid, HkNum.ToString(), commDB, _connection);
                    }
                }
                #endregion

                #region 用户与案件关联

                //string Applicant = dr["申请人"].ToString();

                //if (dr["第二委托人"] != null && !string.IsNullOrEmpty(dr["第二委托人"].ToString()) &&
                //    dr["第二委托人"].ToString().Trim() != "0")
                //{
                //    if (!Applicant.Equals(dr["第二委托人"].ToString()))
                //    {
                //        InsertTCaseClients(dr["第二委托人"].ToString(), HkNum, _rowid, "国内-专利数据补充导入-第二委托人", commDB, _connection);
                //    }
                //}
                //if (dr["第三委托人"] != null && !string.IsNullOrEmpty(dr["第三委托人"].ToString()) &&
                //    dr["第三委托人"].ToString().Trim() != "0")
                //{
                //    if (!Applicant.Equals(dr["第三委托人"].ToString()))
                //    {
                //        InsertTCaseClients(dr["第三委托人"].ToString(), HkNum, _rowid, "国内-专利数据补充导入-第三委托人", commDB, _connection);
                //    }
                //}
                //if (dr["第四委托人"] != null && !string.IsNullOrEmpty(dr["第四委托人"].ToString()) &&
                //    dr["第四委托人"].ToString().Trim() != "0")
                //{
                //    if (!Applicant.Equals(dr["第四委托人"].ToString()))
                //    {
                //        InsertTCaseClients(dr["第四委托人"].ToString(), HkNum, _rowid, "国内-专利数据补充导入-第四委托人", commDB, _connection);
                //    }
                //}
                //if (dr["第五委托人"] != null && !string.IsNullOrEmpty(dr["第五委托人"].ToString()) &&
                //    dr["第五委托人"].ToString().Trim() != "0")
                //{
                //    if (!Applicant.Equals(dr["第五委托人"].ToString()))
                //    {
                //        InsertTCaseClients(dr["第五委托人"].ToString(), HkNum, _rowid, "国内-专利数据补充导入-第五委托人", commDB, _connection);
                //    }
                //}

                #endregion

                #region 增加案件处理人信息

                if (dr["翻译人"] != null && !string.IsNullOrEmpty(dr["翻译人"].ToString()) && dr["翻译人"].ToString().Trim() != "0")
                {
                    InsertUser(dr["翻译人"].ToString(), HkNum, _rowid, "代理部-新申请阶段-翻译人", "国内-专利数据补充导入-翻译人", commDB, _connection);
                }
                if (dr["一校"] != null && !string.IsNullOrEmpty(dr["一校"].ToString()) && dr["一校"].ToString().Trim() != "0")
                {
                    InsertUser(dr["一校"].ToString(), HkNum, _rowid, "代理部-新申请阶段--一校", "国内-专利数据补充导入-一校", commDB, _connection);
                    InsertUser(dr["一校"].ToString(), HkNum, _rowid, "代理部-新申请阶段-办案人", "国内-专利数据补充导入-一校", commDB, _connection);
                }
                if (dr["二校"] != null && !string.IsNullOrEmpty(dr["二校"].ToString()) && dr["二校"].ToString().Trim() != "0")
                {
                   InsertPatentedMemo(HkNum, dr["二校"].ToString(), "二校", _rowid, commDB, _connection);
                }
                string ForeignAgentCode = dr["对外代理人代码"].ToString();
                string ActualAgentCode = dr["实际代理人代码"].ToString();
                if (!string.IsNullOrEmpty(ActualAgentCode) && !string.IsNullOrEmpty(ForeignAgentCode))
                {
                    if (dr["实际代理人代码"] != null && !string.IsNullOrEmpty(dr["实际代理人代码"].ToString()) &&  dr["实际代理人代码"].ToString().Trim() != "0")
                    {
                        InsertUser(dr["实际代理人代码"].ToString(), HkNum, _rowid, "代理部-新申请阶段-办案人", "国内-专利数据补充导入-实际代理人代码", commDB, _connection);
                    }
                }
                else
                {
                    if (dr["实际代理人代码"] != null && !string.IsNullOrEmpty(dr["实际代理人代码"].ToString()) && dr["实际代理人代码"].ToString().Trim() != "0")
                    {
                        InsertUser(dr["实际代理人代码"].ToString(), HkNum, _rowid, "代理部-新申请阶段-办案人", "国内-专利数据补充导入-实际代理人代码", commDB, _connection);
                    }
                    if (dr["对外代理人代码"] != null && !string.IsNullOrEmpty(dr["对外代理人代码"].ToString()) &&
                        dr["对外代理人代码"].ToString().Trim() != "0")
                    {
                        InsertUser(dr["对外代理人代码"].ToString(), HkNum, _rowid, "对外代理人", "国内-专利数据补充导入-对外代理人代码", commDB, _connection);
                    }
                } 
                #endregion

                Result = 1;
            }
            return Result;
        }

        #region 备注
        public int InsertPatentedMemo(int nCaseID, string sContent, string stype, int row, string commDB, SqlConnection _connection)
        {
            int numHk = nCaseID;
            string Sql = "SELECT n_ID  FROM TCase_Memo where  s_Memo='" + sContent.Replace("'", "''") +
                         "' and n_CaseID=" + numHk + " and s_Type='" + stype.Replace("'", "''") + "'";
            int n_ID = _dbHelper.GetbySql(Sql, commDB, _connection);
            int ResultNum = 0;
            if (n_ID <= 0)
            {
                if (!string.IsNullOrEmpty(sContent) ||
                    !string.IsNullOrEmpty(stype))
                {
                    Sql = "INSERT INTO  dbo.TCase_Memo(n_CaseID ,s_Memo ,s_Type ,dt_CreateDate ,dt_EditDate)" +
                          "VALUES  (" + numHk + ",'" + sContent.Replace("'", "''") + "','" +
                          stype.Replace("'", "''") + "','" + DateTime.Now +
                          "','" + DateTime.Now + "')";
                    ResultNum = _dbHelper.InsertbySql(Sql, row, commDB, _connection);
                    if (ResultNum.Equals(0))
                    {
                        _dbHelper.InsertLog(numHk, nCaseID.ToString(), row, "国内-专利数据补充导入-备注", "国内-专利数据补充导入-备注-" + row,
                                            "增加备注失败" + nCaseID.ToString(), Sql.Replace("'", "''"), commDB, _connection);
                    }
                }
            }
            return ResultNum;
        }

        #endregion

        //增加GA01任务链
        private void InsertCaseE(int nID, int _rowid, string nCaseID, string commDB, SqlConnection _connection)
        {
            const string sql = "select * from TCode_Demand where s_sysdemand='GA01'";
            DataTable newTable = _dbHelper.GetDataTablebySql(sql, _connection);
            for (int k = 0; k < newTable.Rows.Count; k++)
            {
                string s_IPType = newTable.Rows[k]["s_IPtype"].ToString();
                string title = newTable.Rows[k]["s_Title"].ToString();
                string description = newTable.Rows[k]["s_Description"].ToString();
                string sCreator = newTable.Rows[k]["s_Creator"].ToString();
                string s_sysDemand = newTable.Rows[k]["s_sysDemand"].ToString();
                string n_DemandType = newTable.Rows[k]["n_ID"].ToString();

                string Sql = "INSERT INTO dbo.T_Demand(s_sourcetype1,s_ModuleType,s_Title,s_Description,s_Creator,s_Editor,s_IPType,s_SysDemand,n_DemandType,dt_EditDate,dt_CreateDate,s_SourceModuleType,n_CaseID)" +
                             "VALUES  ('国内-专利数据补充导入--客户指示提实审" + _rowid + "','Case','" + title + "','" + description + "','" + sCreator + "','" + sCreator + "','" + s_IPType + "','" + s_sysDemand + "','" + n_DemandType + "','" + DateTime.Now + "','" + DateTime.Now + "','Case','" + nCaseID + "')";

                //查询是否存在此案件要求要求
                string strSql = "select n_ID from T_Demand where s_ModuleType='Case'  and n_CaseID=" + nCaseID + " and  s_sourcetype1='国内-专利数据补充导入--客户指示提实审" + _rowid + "' and s_Title='" + title + "'";
                DataTable Table = _dbHelper.GetDataTablebySql(strSql, _connection);

                if (Table.Rows.Count <= 0)
                {
                    _dbHelper.InsertbySql(Sql, _rowid, commDB, _connection);
                }
            }
        } 

        //增加要求
        private void InsertPatentSemand(string s_Title, int n_CaseID, int _rowid, string commDB, SqlConnection _connection)
        {
            string Sql = "SELECT n_ID  FROM T_Demand where  s_Title='" + s_Title.ToString().Replace("'", "''") +
                         "' and n_CaseID=" + n_CaseID;
            int n_ID = _dbHelper.GetbySql(Sql, commDB, _connection);
            if (n_ID <= 0)
            {
                Sql = "INSERT INTO dbo.T_Demand(s_sourcetype1,dt_EditDate,s_Title,dt_CreateDate,n_CaseID)" +
                      "VALUES  ('国内-专利数据补充导入-年费备注','" + DateTime.Now + "','" + s_Title.ToString().Replace("'", "''") + "','" + DateTime.Now +
                      "'," + n_CaseID + ")";
                int ResultNum = _dbHelper.InsertbySql(Sql, _rowid, commDB, _connection);
                if (ResultNum.Equals(0))
                {
                    _dbHelper.InsertLog(n_CaseID, "", _rowid, "国内-专利数据补充导入", "国内-专利数据补充导入-增加年费要求" + _rowid, "增加年费要求失败", Sql.Replace("'", "''"), commDB, _connection);
                }
            } 
        }

        //增加用户与案件相关联  【相关客户】
        public void InsertTCaseClients(string s_ClientCode, int n_CaseID, int _rowid, string tableName, string commDB, SqlConnection _connection)
        {
            //查询人员ID
            string Sql = " SELECT n_ClientID FROM TCstmr_Client WHERE s_ClientCode='" + s_ClientCode + "'";
            int n_ClientID = _dbHelper.GetbySql(Sql, commDB, _connection);
            if (n_ClientID > 0)
            {
                //判断是否存在此人
                string strSql = "SELECT count(*) sunmCount FROM TCase_Clients WHERE n_CaseID=" + n_CaseID +
                                " AND n_ClientID=" + n_ClientID;
                if (_dbHelper.GetbySql(strSql, commDB, _connection) <= 0)
                {

                    Sql = "INSERT INTO dbo.TCase_Clients( n_CaseID, n_ClientID)" +
                          "VALUES  (" + n_CaseID + " , " + n_ClientID + ")";

                    if (_dbHelper.InsertbySql(Sql, _rowid, commDB, _connection) <= 0)
                    {
                        _dbHelper.InsertLog(n_CaseID, "", _rowid, tableName, tableName + _rowid, "增加用户与案件关系失败,用户编号为：" + s_ClientCode, Sql.Replace("'", "''"), commDB, _connection);
                    } 
                }
                else
                {
                    _dbHelper.InsertLog(n_CaseID, "", _rowid, tableName, tableName + _rowid, "案件未存在此用户的相关客户信息", strSql.Replace("'", "''"), commDB, _connection);
                }
            }
            else
            {
                _dbHelper.InsertLog(n_CaseID, "", _rowid, tableName, tableName + _rowid, "未查询到用户信息,用户编号为：" + s_ClientCode, Sql.Replace("'", "''"), commDB, _connection);
            }
        }
        
        //增加案件处理人
        public void InsertUser(string s_InternalCode, int n_CaseID, int _rowid, string NameType, string tableName, string commDB, SqlConnection _connection)
        {
            //查询人员ID
            string Sql = "SELECT n_ID FROM dbo.TCode_Employee WHERE s_InternalCode='" + s_InternalCode +
                         "'  OR s_Name='" + s_InternalCode + "'";
            int n_AssignorID = _dbHelper.GetbySql(Sql, commDB, _connection);

            //案件角色
            string Sql2 = "SELECT n_ID FROM dbo.TCode_CaseRole WHERE s_Name LIKE '%" + NameType + "%'";
            int n_CaseRoleID = _dbHelper.GetbySql(Sql2, commDB, _connection);

            if (n_AssignorID > 0)
            {
                Sql = "SELECT COUNT(*) AS sumCount FROM dbo.TCase_Attorney WHERE n_CaseID=" + n_CaseID +
                      " AND n_AttorneyID=" + n_AssignorID + " and n_CaseRoleID=" + n_CaseRoleID;
                if (_dbHelper.GetbySql(Sql, commDB, _connection) <= 0)
                {
                    Sql =
                        "INSERT INTO dbo.TCase_Attorney( n_CaseID ,dt_AssignDate,n_AssignorID,n_AttorneyID,n_CaseRoleID)" +
                        "VALUES  (" + n_CaseID + " , '" + DateTime.Now + "',1000131," + n_AssignorID + "," +
                        n_CaseRoleID + ")";
                    _dbHelper.InsertbySql(Sql, _rowid, commDB, _connection);
                } 
            }
            else
            {
                if (!NameType.Equals("代理部-新申请阶段-办案人"))
                {
                    _dbHelper.InsertLog(n_CaseID, "", _rowid, "国内-专利数据补充导入", "国内-专利数据补充导入-增加案件处理人" + _rowid, "未找到人员为：" + s_InternalCode, Sql.Replace("'", "''"), commDB, _connection);
                } 
            }
        }

        #endregion 

        #region  香港-专利数据补充导入
        public int HongKang(int rowid, DataRow dr, int OutFileID, string commDB, SqlConnection _connection)
        {
            int result = 0;
            string sNo = dr["HK卷号"].ToString().Trim();
            int hkNum = _dbHelper.GetIDbyName(sNo, 4,_connection);
            if (hkNum.Equals(0))
            {
                _dbHelper.InsertLog(0, sNo, rowid, "香港-专利数据补充导入", "香港-专利数据补充导入-" + rowid, "未找到“我方卷号”为：" + sNo, "", commDB, _connection);
                return 0;
            }
            if (hkNum != 0)
            {
                #region 香港表
                string sql = "SELECT * FROM TCode_BusinessType WHERE s_Code IN ('HT','HS','HD') AND n_ID IN (SELECT n_BusinessTypeID  FROM dbo.TCase_Base WHERE n_CaseID=" +
                        hkNum + " AND n_CaseID IN ( SELECT n_CaseID FROM dbo.TPCase_Patent))";
                DataTable tablew = _dbHelper.GetDataTablebySql(sql,_connection);
                if (tablew.Rows.Count > 0)
                {
                    #region
                    string strSql =
                        " UPDATE TCase_Base SET ObjectType=(SELECT oid FROM dbo.XPObjectType WHERE TypeName='DataEntities.Case.Patents.HongKongApplication') WHERE n_CaseID=" +
                        hkNum;
                    _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                    DataTable tableCount =_dbHelper.GetDataTablebySql("SELECT COUNT(*) as sumCount FROM TPCase_HongKongApplication  WHERE n_CaseID=" + hkNum, _connection);
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
                        using (SqlCommand cmd = _connection.CreateCommand())
                        {
                            cmd.CommandText = strSql;
                            result = cmd.ExecuteNonQuery();
                            if (result == 0)
                            {
                                _dbHelper.InsertLog(hkNum, sNo, rowid, "香港-专利数据补充导入", "香港-专利数据补充导入-" + rowid, "未找到“我方卷号”为：" + sNo, strSql.Replace("'", "''"), commDB, _connection);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _dbHelper.InsertLog(hkNum, sNo, rowid, "香港-专利数据补充导入", "香港-专利数据补充导入-" + rowid, "香港-专利数据补充导入错误信息：" + ex.Message, strSql.Replace("'", "''"), commDB, _connection);
                    }
                }
                else
                {
                    _dbHelper.InsertLog(hkNum, sNo, rowid, "香港-专利数据补充导入", "香港-专利数据补充导入-" + rowid, "案件不存在香港关系", sql.Replace("'", "''"), commDB, _connection);
                }

                #endregion

                #region 原案信息表
                string sqlO = " SELECT n_OrigPatInfoID  FROM dbo.TPCase_Patent where n_CaseID=" + hkNum;
                int Country = _dbHelper.GetIDbyName(dr["母案国家"].ToString(), 1,_connection);
                int InNum = _dbHelper.GetbySql(sqlO, commDB, _connection);
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
                if (_dbHelper.InsertbySql(sqlO, rowid, commDB, _connection) == 0)
                {
                    _dbHelper.InsertLog(hkNum, sNo, rowid, "香港-专利数据补充导入", "香港-专利数据补充导入-" + rowid, "香港-专利数据补充'原案信息'导入错误信息", sqlO.Replace("'", "''"), commDB, _connection);
                }
                else
                {
                    sqlO = "select n_ID from TPCase_OrigPatInfo where s_CaseSerial='" +
                           dr["母案卷号"].ToString().Replace("'", "''") + "' and s_AppNo='" +
                           dr["母案申请号"].ToString().Replace("'", "''") + "' and n_OrigRegCountry=" + Country;
                    int SQLID = _dbHelper.GetbySql(sqlO, commDB, _connection);
                    if (SQLID > 0)
                    {
                        sqlO = "update dbo.TPCase_Patent set n_OrigPatInfoID=" + SQLID + " where n_CaseID=" + hkNum;
                        _dbHelper.InsertbySql(sqlO, rowid, commDB, _connection);
                    }
                    else
                    {
                        _dbHelper.InsertLog(hkNum, sNo, rowid, "香港-专利数据补充导入", "香港-专利数据补充导入-" + rowid, "原案信息不存在", sqlO.Replace("'", "''"), commDB, _connection);
                    }
                }

                #endregion

                #region 案件关系

                int hkNumB = _dbHelper.GetIDbyName(dr["母案卷号"].ToString().Trim(), 4,_connection);
                T_MainFiles _tMainFiles;
                if (hkNum > 0 && hkNumB > 0)
                {
                    const string strSql = "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='香港申请' AND s_MasterName='国内母案' AND s_SlaveName='香港案' AND s_IPType='P'";
                    //案件关系配置表
                    int n_ID = _dbHelper.GetbySql(strSql, commDB, _connection);
                    int NUMS =
                        _dbHelper.GetbySql("SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where n_CaseIDA=" + hkNum +
                                   " and n_CaseIDB=" + hkNumB + " and n_CodeRelativeID=" + n_ID, commDB, _connection);
                    if (NUMS <= 0 && hkNum > 0)
                    {
                        _tMainFiles = new T_MainFiles();
                        _tMainFiles.InsertTCaseCaseRelative(hkNum, hkNumB, n_ID, 0, rowid, commDB, _connection);
                    }
                    else
                    {
                        _dbHelper.InsertLog(hkNum, sNo, rowid, "香港-专利数据补充导入", "香港-专利数据补充导入-" + rowid, "缺少案件ID，无法建立关系", "", commDB, _connection);
                    }
                }

                #endregion

                #region 更新外代理人信息
                string s_CoopAgencyToNo = dr["外代理名称"].ToString().Trim();
                string strSqlAgency = "select n_AgencyID from TCstmr_CoopAgency where s_Name='" + s_CoopAgencyToNo + "' or s_NativeName='" + s_CoopAgencyToNo + "'";
                int Num = _dbHelper.GetbySql(strSqlAgency, commDB, _connection);

                if (Num > 0)
                {
                    strSqlAgency = " UPDATE TCase_Base SET s_CoopAgencyToNo='" + dr["外代理文号"].ToString().Trim() + "',n_CoopAgencyToID=" + Num + " WHERE n_CaseID=" + hkNum;
                    _dbHelper.InsertbySql(strSqlAgency, 0, commDB, _connection);
                }
                else
                {
                    _dbHelper.InsertLog(hkNum, sNo, rowid, "香港-专利数据补充导入", "香港-专利数据补充导入-" + rowid, "无外代理信息,外代理名称：" + s_CoopAgencyToNo, "", commDB, _connection);
                }
                #endregion

                #region 生成发文记录
                int caseNo = 0;
                int agent = 0;
                string StrSql = "select n_ClientID,n_CoopAgencyToID from tcase_base where s_caseSerial='" + sNo + "'";
                DataTable table = _dbHelper.GetDataTablebySql(StrSql, _connection);
                if (table.Rows.Count > 0)
                {
                    caseNo = table.Rows[0]["n_ClientID"].ToString() == "" ? 0 : int.Parse(table.Rows[0]["n_ClientID"].ToString());
                    agent = table.Rows[0]["n_CoopAgencyToID"].ToString() == "" ? 0 : int.Parse(table.Rows[0]["n_CoopAgencyToID"].ToString());
                }
                //s_IOType I：来文 O：发文 T：其它文件 
                //s_ClientGov  C: 客户  O: 官方 
                string sClientGov = "";
                #region 发官方文 
                _tMainFiles = new T_MainFiles();
                //根据数据表中提供的第一步P4寄出日建立一个发外代文件，日期作为发文日；
                string One = dr["第一步P4寄出日"].ToString().Trim() == ""
                                                     ? string.Empty
                                                     : "'" + dr["第一步P4寄出日"] + "'";
                if (!string.IsNullOrEmpty(One))
                {
                    sClientGov = "O";
                    _tMainFiles.InTMainFiles("第一步P4寄出日", One, hkNum, rowid, agent, 0, sClientGov, OutFileID, commDB, _connection);
                }
                //根据数据表中提供的第二步P5寄出日建立一个发外代文件，日期作为发文日；
                string Two = dr["第二步P5寄出日"].ToString().Trim() == ""
                                                    ? string.Empty
                                                    : "'" + dr["第二步P5寄出日"] + "'";
                if (!string.IsNullOrEmpty(Two))
                {
                    sClientGov = "O";
                    _tMainFiles.InTMainFiles("第二步P5寄出日", Two, hkNum, rowid, agent, 0, sClientGov, OutFileID, commDB, _connection);
                }
                #endregion

                #region 发客户文
                string Three = dr["第一步转公开"].ToString().Trim() == ""
                                                   ? string.Empty
                                                   : "'" + dr["第一步转公开"] + "'";
                if (!string.IsNullOrEmpty(Three))
                {
                    sClientGov = "C";
                    _tMainFiles.InTMainFiles("第一步转公开", Three, hkNum, rowid, 0, caseNo, sClientGov, OutFileID, commDB, _connection);
                }
                string Four = dr["第二步转授权"].ToString().Trim() == ""
                                                   ? string.Empty
                                                   : "'" + dr["第二步转授权"] + "'";
                if (!string.IsNullOrEmpty(Four))
                {
                    sClientGov = "C";
                    _tMainFiles.InTMainFiles("第二步转授权", Four, hkNum, rowid, 0, caseNo, sClientGov, OutFileID, commDB, _connection);
                }
                string dt_stAgentReports = dr["第一步代理报告"].ToString().Trim() == ""
                                                                  ? string.Empty
                                                                  : "'" + dr["第一步代理报告"] + "'";
                if (!string.IsNullOrEmpty(dt_stAgentReports))
                {
                    sClientGov = "C";
                    _tMainFiles.InTMainFiles("第一步代理报告", dt_stAgentReports, hkNum, rowid, 0, caseNo, sClientGov, OutFileID, commDB, _connection);
                }
                string dt_2ndAgentReports = dr["第二步代理报告"].ToString().Trim() == ""
                                                                  ? string.Empty
                                                                  : "'" + dr["第二步代理报告"] + "'";
                if (!string.IsNullOrEmpty(dt_2ndAgentReports))
                {
                    sClientGov = "C";
                    _tMainFiles.InTMainFiles("第二步代理报告", dt_2ndAgentReports, hkNum, rowid, 0, caseNo, sClientGov, OutFileID, commDB, _connection);
                }
                #endregion
                #endregion

                result = 1;
            }
            return result;
        }
        #endregion 

        #region 分案信息
        public int UpdateCaseDivisionInfo(DataRow row, int rowid, string commDB, SqlConnection _connection)
        {
            int result = 0;
            string sNo = row["我方卷号"].ToString();
            int HkNum = _dbHelper.GetIDbyName(sNo, 2, _connection);

            if (HkNum.Equals(0))
            {
                _dbHelper.InsertLog(0, sNo, rowid, "分案信息", "分案信息-" + rowid, "未找到“我方卷号”为：" + sNo, "", commDB, _connection);
                return result;
            }
            else
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
                    _dbHelper.InsertbySql(Sql, rowid, commDB, _connection); 
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
                        DataTable table = GetsCaseSerial(s2,_connection);
                        string Master = "";
                        if (table.Rows.Count > 0)
                        {
                            for (int i = 0; i < table.Rows.Count; i++)
                            {
                                string sCaseSerial = table.Rows[i]["s_CaseSerial"].ToString();
                                string[] arry = sCaseSerial.Split('-');
                                if (!sCaseSerial.Equals(sNo))
                                {
                                    if (arry[0].Equals(s2))
                                    {
                                        Master = sCaseSerial;
                                        InsertCodeCaseRelative(HkNum, sCaseSerial, "母案", rowid, commDB, _connection);
                                    }
                                    else
                                    {
                                        InsertCodeCaseRelative(HkNum, sCaseSerial, "相关案件", rowid, commDB, _connection);
                                        if (!string.IsNullOrEmpty(Master))
                                        {
                                            InsertCodeCaseRelative(_dbHelper.GetIDbyName(Master, 2, _connection), sCaseSerial, "分案", rowid, commDB, _connection);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(row["针对分案文号"].ToString()))
                    {
                        InsertCodeCaseRelative(HkNum, row["针对分案文号"].ToString(), "相关案件", rowid, commDB, _connection);
                    }
                    result = 1;
                }
                catch (Exception ex)
                {
                    _dbHelper.InsertLog(HkNum, sNo, rowid, "分案信息", "分案信息-" + rowid, "更新分案申请信息错误" + ex.Message, "", commDB, _connection);
               
                }
            } 
            return result;
        }
 
        //模糊查询
        private DataTable GetsCaseSerial(string sCaseSerial, SqlConnection _connection)
        {
            var table = new DataTable();
            table.Clear();
            string sql = "SELECT s_CaseSerial FROM dbo.TCase_Base WHERE s_CaseSerial LIKE '" + sCaseSerial + "%'";
            table = _dbHelper.GetDataTablebySql(sql, _connection);
            return table;
        }

        //增加分案信息
        private void InsertCodeCaseRelative(int HKNum, string sOrigCaseNo, string Type, int rowid, string commDB, SqlConnection _connection)
        {
            var _tMainFiles = new T_MainFiles(); 
            //查询母案 
            int caseID = _dbHelper.GetIDbyName(sOrigCaseNo, 2, _connection);

            string strSql =
                   "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='分案' AND s_MasterName='母案' AND s_SlaveName='子案' AND s_IPType='P'";
            if (Type.Equals("相关案件"))
            {
                strSql =
                    "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='相关案件' AND s_MasterName='' AND s_SlaveName='' AND s_IPType='P'";
            }
            int n_ID = _dbHelper.GetbySql(strSql, commDB, _connection);
            if (caseID > 0)
            {
                int NUMS =
                    _dbHelper.GetbySql("SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where ((n_CaseIDA=" + HKNum +
                               " and n_CaseIDB=" + caseID + ") or  (n_CaseIDA=" + caseID + " and n_CaseIDB=" + HKNum + "))and n_CodeRelativeID=" + n_ID, commDB, _connection);
                if (NUMS <= 0 && HKNum > 0)
                {
                    if (Type.Equals("母案") || Type.Equals("相关案件"))
                    {
                        _tMainFiles.InsertTCaseCaseRelative(HKNum, caseID, n_ID, 0, rowid, commDB, _connection);
                    }
                    else
                    {
                        _tMainFiles.InsertTCaseCaseRelative(caseID, HKNum, n_ID, 0, rowid, commDB, _connection);
                    }
                }
            }
        }
      
        #endregion 

        #region 最后一步更新
        public int UpdateCase(DataRow dataRow, int rowid, string commDB, SqlConnection _connection)
        {
            //我方文号	官方中文名称	专利名称外文	申请号	申请日	PCT申请号	PCT申请日	公开号	公开日	授权公告号	授权公告日	证书号	发证书日	
            //申请人中文名称	申请人英文名称	发明人	发明人英文名称	第一代理人	第二代理人
            string s_No = dataRow["我方卷号"].ToString().Trim();
            int HkNum = _dbHelper.GetIDbyName(s_No, 2, _connection);
            if (HkNum.Equals(0))
            {
                _dbHelper.InsertLog(0, s_No, rowid, "Update", "Update-" + rowid, "未找到“我方卷号”为：" + s_No, "", commDB, _connection);
                return 0;
            }
            else
            {
                #region 更新案件信息
                string strSql = "update TCase_Base ";
                if (dataRow["专利名称外文"] != null && !string.IsNullOrEmpty(dataRow["专利名称外文"].ToString().Trim()))
                {
                    if (strSql.Contains("set"))
                    {
                        strSql += ",s_CaseOtherName='" + dataRow["专利名称外文"].ToString().Trim() + "'";
                    }
                    else
                    {
                        strSql += " set s_CaseOtherName='" + dataRow["专利名称外文"].ToString().Trim() + "'";
                    }
                }
                if (dataRow["官方中文名称"] != null && !string.IsNullOrEmpty(dataRow["官方中文名称"].ToString().Trim()))
                {
                    if (strSql.Contains("set"))
                    {
                        strSql += ",s_ClientCName='" + dataRow["官方中文名称"].ToString().Trim() + "'";
                    }
                    else
                    {
                        strSql += " set s_ClientCName='" + dataRow["官方中文名称"].ToString().Trim() + "'";
                    }
                }
                if (dataRow["第一代理人"] != null && !string.IsNullOrEmpty(dataRow["第一代理人"].ToString().Trim()))
                {
                    if (strSql.Contains("set"))
                    {
                        strSql += ",n_FirstAttorney='" + dataRow["第一代理人"].ToString().Trim() + "'";
                    }
                    else
                    {
                        strSql += " set n_FirstAttorney='" + dataRow["第一代理人"].ToString().Trim() + "'";
                    }
                }
                if (dataRow["第二代理人"] != null && !string.IsNullOrEmpty(dataRow["第二代理人"].ToString().Trim()))
                {
                    if (strSql.Contains("set"))
                    {
                        strSql += ",n_SecondAttorney='" + dataRow["第二代理人"].ToString().Trim() + "'";
                    }
                    else
                    {
                        strSql += " set n_SecondAttorney='" + dataRow["第二代理人"].ToString().Trim() + "'";
                    }
                }
                strSql += " where n_CaseID=" + HkNum;
                if (_dbHelper.InsertbySql(strSql, rowid, commDB, _connection) <= 0)
                {
                    _dbHelper.InsertLog(HkNum, s_No, rowid, "Update", "Update-" + rowid, "更新案件信息失败",
                                        strSql.Replace("'", "''"), commDB, _connection);
                } 
                #endregion 

                #region 更新法律信息
                strSql = "Update dbo.TPCase_LawInfo ";
                if (dataRow["申请号"] != null && !string.IsNullOrEmpty(dataRow["申请号"].ToString().Trim()))
                {
                    if (strSql.Contains("set"))
                    {
                        strSql += " s_AppNo='" + dataRow["申请号"].ToString().Trim() + "'";
                    }
                    else
                    {
                        strSql += " set s_AppNo='" + dataRow["申请号"].ToString().Trim() + "'";
                    }
                }
                if (dataRow["申请日"] != null && !string.IsNullOrEmpty(dataRow["申请日"].ToString().Trim()))
                {
                    if (strSql.Contains("set"))
                    {
                        strSql += ",dt_AppDate='" + dataRow["申请日"].ToString().Trim() + "'";
                    }
                    else
                    {
                        strSql += " set dt_AppDate='" + dataRow["申请日"].ToString().Trim() + "'";
                    }
                }
                if (dataRow["PCT申请号"] != null && !string.IsNullOrEmpty(dataRow["PCT申请号"].ToString().Trim()))
                {
                    if (strSql.Contains("set"))
                    {
                        strSql += ",s_PCTAppNo='" + dataRow["PCT申请号"].ToString().Trim() + "'";
                    }else
                    {
                        strSql += "set s_PCTAppNo='" + dataRow["PCT申请号"].ToString().Trim() + "'"; 
                    }
                }
                if (dataRow["PCT申请日"] != null && !string.IsNullOrEmpty(dataRow["PCT申请日"].ToString().Trim()))
                {
                    if (strSql.Contains("set"))
                    {
                        strSql += ",dt_PCTAppDate='" + dataRow["PCT申请日"].ToString().Trim() + "'";
                    }
                    else
                    {
                        strSql += " set dt_PCTAppDate='" + dataRow["PCT申请日"].ToString().Trim() + "'";
                    }
                }
                if (dataRow["公开号"] != null && !string.IsNullOrEmpty(dataRow["公开号"].ToString().Trim()))
                {
                    if (strSql.Contains("set"))
                    {
                        strSql += ",s_PubNo='" + dataRow["公开号"].ToString().Trim() + "'";
                    }
                    else
                    {
                        strSql += " set s_PubNo='" + dataRow["公开号"].ToString().Trim() + "'";
                    }
                }
                if (dataRow["公开日"] != null && !string.IsNullOrEmpty(dataRow["公开日"].ToString().Trim()))
                {
                    if (strSql.Contains("set"))
                    {
                        strSql += ",dt_PubDate='" + dataRow["公开日"].ToString().Trim() + "'";
                    }
                    else
                    {
                        strSql += " set dt_PubDate='" + dataRow["公开日"].ToString().Trim() + "'";
                    }
                }
                if (dataRow["授权公告号"] != null && !string.IsNullOrEmpty(dataRow["授权公告号"].ToString().Trim()))
                {
                    if (strSql.Contains("set"))
                    {
                        strSql += ",s_IssuedPubNo='" + dataRow["授权公告号"].ToString().Trim() + "'";
                    }
                    else
                    {
                        strSql += " set s_IssuedPubNo='" + dataRow["授权公告号"].ToString().Trim() + "'";
                    }
                }
                if (dataRow["授权公告日"] != null && !string.IsNullOrEmpty(dataRow["授权公告日"].ToString().Trim()))
                {
                    if (strSql.Contains("set"))
                    {
                        strSql += ",dt_IssuedPubDate='" + dataRow["授权公告日"].ToString().Trim() + "'";
                    }
                    else
                    {
                        strSql += " set dt_IssuedPubDate='" + dataRow["授权公告日"].ToString().Trim() + "'";
                    }
                }
                if (dataRow["证书号"] != null && !string.IsNullOrEmpty(dataRow["证书号"].ToString().Trim()))
                {
                    if (strSql.Contains("set"))
                    {
                        strSql += ",s_CertfNo='" + dataRow["证书号"].ToString().Trim() + "'";
                    }
                    else
                    {
                        strSql += " set s_CertfNo='" + dataRow["证书号"].ToString().Trim() + "'";
                    }
                }
                if (dataRow["发证书日"] != null && !string.IsNullOrEmpty(dataRow["发证书日"].ToString().Trim()))
                {
                    if (strSql.Contains("set"))
                    {
                        strSql += ",dt_CertfDate='" + dataRow["发证书日"].ToString().Trim() + "'";
                    }
                    else
                    {
                        strSql += "set dt_CertfDate='" + dataRow["发证书日"].ToString().Trim() + "'";
                    }
                } 
                strSql+=" WHERE n_ID IN (SELECT n_LawID FROM dbo.TPCase_Patent WHERE n_CaseID=" + HkNum + ")";
                if (_dbHelper.InsertbySql(strSql, rowid, commDB, _connection) <= 0)
                {
                    _dbHelper.InsertLog(HkNum, s_No, rowid, "Update", "Update-" + rowid, "更新法律信息失败",
                                        strSql.Replace("'", "''"), commDB, _connection);
                }
                else
                {
                    return 1;
                } 
                #endregion 

                #region 更新申请人 
                var sAppName = new string[] {};
                if(dataRow["申请人中文名称"] != null && !string.IsNullOrEmpty(dataRow["申请人中文名称"].ToString().Trim()))
                {
                    sAppName = dataRow["申请人中文名称"].ToString().Trim().Split(';');
                }
                var sEAppName = new string[] {};
                if (dataRow["申请人英文名称"] != null && !string.IsNullOrEmpty(dataRow["申请人英文名称"].ToString().Trim()))
                {
                    sEAppName = dataRow["申请人英文名称"].ToString().Trim().Split(';');
                }
                //查找案件申请人数量
                strSql = "SELECT COUNT(*) AS suncount FROM TCase_Applicant WHERE n_CaseID="+HkNum;
                int Count = _dbHelper.GetbySql(strSql, commDB, _connection);
                if (sAppName.Length.Equals(sEAppName.Length) && sAppName.Length.Equals(Count))
                {
                    for (var i = 0; i < sAppName.Length; i++)
                    {
                        if(!string.IsNullOrEmpty(sAppName[i].ToString()))
                        {
                            strSql = "SELECT n_AppID FROM dbo.TCstmr_Applicant WHERE s_Name='" + sAppName[i] +
                                     "' AND s_NativeName='" + sEAppName[i] + "'";
                            int nAppID = _dbHelper.GetbySql(strSql, commDB, _connection);
                            if (nAppID <= 0)
                            {
                                _dbHelper.InsertLog(0, s_No, rowid, "Update", "Update-申请人-" + rowid + "第" + i + "申请人", "未查询到申请人信息", strSql.Replace("'", "''"), commDB, _connection);
                            }
                            else
                            {
                                strSql = "UPDATE  TCase_Applicant SET s_Name='" + sAppName[i] + "',s_NativeName='" + sEAppName[i] + "'  WHERE n_CaseID="+HkNum+" AND n_ApplicantID=" + nAppID;
                                _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                            }
                        }
                        else
                        {
                            _dbHelper.InsertLog(0, s_No, rowid, "Update", "Update-申请人-" + rowid + "第" + i + "申请人", "申请人名称为空", "", commDB, _connection);
                        }
                    }
                }
                else
                {
                    _dbHelper.InsertLog(0, s_No, rowid, "Update", "Update-申请人-" + rowid, "申请人数量和案件申请人数量不相等" + dataRow["申请人中文名称"].ToString().Trim(), "", commDB, _connection);
                }
                #endregion

                #region 更新发明人
                var sInventor = new string[] { };
                if (dataRow["发明人"] != null && !string.IsNullOrEmpty(dataRow["发明人"].ToString().Trim()))
                {
                    sInventor = dataRow["发明人"].ToString().Trim().Split(';');
                }
                var sEInventor = new string[] { };
                if (dataRow["发明人英文名称"] != null && !string.IsNullOrEmpty(dataRow["发明人英文名称"].ToString().Trim()))
                {
                    sEInventor = dataRow["发明人英文名称"].ToString().Trim().Split(';');
                }
                //查找案件申请人数量
                strSql = "SELECT n_ID FROM TPCase_Inventor WHERE n_CaseID=" + HkNum + " ORDER BY n_Sequence ASC ";
                DataTable _table = _dbHelper.GetDataTablebySql(strSql, _connection);
                if (sInventor.Length.Equals(sEInventor.Length) && sInventor.Length.Equals(_table.Rows.Count))
                {
                    for (var i = 0; i < sInventor.Length; i++)
                    {
                        if (!string.IsNullOrEmpty(sInventor[i].ToString()))
                        {
                            strSql = "SELECT n_ID FROM dbo.TPCase_Inventor WHERE n_ID=" + _table.Rows[i]["n_ID"];
                            int nID = _dbHelper.GetbySql(strSql, commDB, _connection);
                            if (nID <= 0)
                            {
                                _dbHelper.InsertLog(0, s_No, rowid, "Update", "Update-发明人-" + rowid + "第" + i + "发明人", "查找发明人错误", strSql.Replace("'", "''"), commDB, _connection);
                            }
                            else
                            {
                                strSql = "UPDATE  TPCase_Inventor SET s_Name='" + sAppName[i] + "',s_NativeName='" + sEAppName[i] + "'  WHERE n_ID=" + nID;
                                _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                            }
                        }
                        else
                        {
                            _dbHelper.InsertLog(0, s_No, rowid, "Update", "Update-发明人-" + rowid + "第" + i + "发明人", "发明人名称为空", "", commDB, _connection);
                        }
                    }
                }
                else
                {
                    _dbHelper.InsertLog(0, s_No, rowid, "Update", "Update-发明人-" + rowid, "申请人数量和案件发明人数量不相等" + dataRow["发明人"].ToString().Trim(), "", commDB, _connection);
                }
                #endregion

            } 
            return 0;
        }

        #endregion 
    }
}
