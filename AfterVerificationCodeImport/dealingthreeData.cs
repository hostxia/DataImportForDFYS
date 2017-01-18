using System;
using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport
{
    class dealingthreeData
    {
        readonly DBHelper _dbHelper = new DBHelper();
        readonly dealingClient _dealingClientData = new dealingClient();

        #region 实体基本信息
        public int TCaseApplicant(int rowid, DataRow dr, string commDB, SqlConnection _connection)
        {
            int Result = 0;
            string strSql = "";
            string sCaserial = dr["我方卷号"].ToString();
            int HkNum = _dbHelper.GetIDbyName(sCaserial, 2,_connection);
            if (HkNum.Equals(0))
            {
                _dbHelper.InsertLog(0, sCaserial, rowid, "国外-实体信息表", "国外-实体信息表-" + rowid, "未找到“我方卷号”为：" + sCaserial, "", commDB, _connection);
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
                        int n_AgencyID = _dbHelper.GetIDbyName(s_NoN, 5,_connection);
                        if (n_AgencyID > 0)
                        {
                            string strSqlA = " update TCase_Base set n_CoopAgencyToID=" + n_AgencyID +
                                             " where n_CaseID=" + HkNum;
                            _dbHelper.InsertbySql(strSqlA, rowid, commDB, _connection);//同步代理机构ID
                            //添加代理机构联系人
                            strSql = "select n_ContactID,n_Sequence  from TCstmr_AgencyContact  where n_AgencyID=" + n_AgencyID;
                            DataTable tableA = _dbHelper.GetDataTablebySql(strSql,_connection);
                            for (int ik = 0; ik < tableA.Rows.Count; ik++)
                            {
                                string InsertInto = " insert into TCase_ToAgencyContact(n_CaseID,n_ContactID,n_Sequence)values(" + HkNum + ",'" + tableA.Rows[ik]["n_ContactID"] + "','" + tableA.Rows[ik]["n_Sequence"] + "')";
                               Result= _dbHelper.InsertbySql(InsertInto, rowid, commDB, _connection);
                            }
                        }
                        else
                        {  Result = 0;
                            _dbHelper.InsertLog(HkNum, sCaserial, rowid, "国外-实体信息表", "国外-实体信息表-" + rowid, "实体ID未查到-外代理：" + s_NoN, "", commDB, _connection);
                        }
                    }
                    else
                    {  Result = 0;
                        _dbHelper.InsertLog(HkNum, sCaserial, rowid, "国外-实体信息表", "国外-实体信息表-" + rowid, "实体ID为空-外代理：" + s_NoN, "", commDB, _connection);
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
                        int numweituoren = _dbHelper.GetbySql(strSql2, commDB, _connection);
                        if (numweituoren > 0)
                        {
                            strSql = "  update TCase_Base set s_ClientSerial='" + dr["实体方卷号"] + "',n_ClientID=" +
                                     numweituoren + " where n_CaseID=" + HkNum;

                            Result = _dbHelper.InsertbySql(strSql, rowid, commDB, _connection); 
                            if (Result == 0)
                            {
                                _dbHelper.InsertLog(HkNum, sCaserial, rowid, "国外-实体信息表", "国外-实体信息表-" + rowid, "实体信息-委托人更新失败，实体方卷号:" + dr["实体方卷号"], strSql.Replace("'", "''"), commDB, _connection);
                            }
                        }
                        else
                        {
                            _dbHelper.InsertLog(HkNum, sCaserial, rowid, "国外-实体信息表", "国外-实体信息表-" + rowid, "实体ID未查到-委托人：" + dr["实体ID"].ToString().Trim(), strSql2, commDB, _connection); 
                              Result = 0;
                        }
                    }
                    else
                    {
                        _dbHelper.InsertLog(HkNum, sCaserial, rowid, "国外-实体信息表", "国外-实体信息表-" + rowid, "实体ID为空-委托人：" + dr["实体ID"].ToString().Trim(), "", commDB, _connection);
                        Result = 0;
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
                        #region 申请人
                        int nAppID = _dbHelper.GetIDbyName(sNo, 3,_connection);
                        if (nAppID > 0)
                        {
                            strSql = "SELECT n_ID FROM dbo.TCase_Applicant  WHERE n_CaseID=" + HkNum +
                                     "  AND n_ApplicantID=" + nAppID;
                            if (_dbHelper.GetbySql(strSql, commDB, _connection) <= 0)
                            {
                                int nApplicantID = _dbHelper.GetbySql("SELECT n_AppID  FROM dbo.TCstmr_Applicant WHERE s_AppCode='" +
                                                                      dr["实体ID"].ToString().Trim().Replace("'", "''") + "'", commDB, _connection);
                                if (nApplicantID > 0)
                                {
                                    string sNativeName =
                                        _dbHelper.GetStringbySql(
                                            "SELECT s_NativeName  FROM dbo.TCstmr_Applicant WHERE s_AppCode='" +
                                            dr["实体ID"].ToString().Trim().Replace("'", "''") + "'", _connection).ToString();
                                    string sName =
                                        _dbHelper.GetStringbySql(
                                            "SELECT s_Name  FROM dbo.TCstmr_Applicant WHERE s_AppCode='" +
                                            dr["实体ID"].ToString().Trim().Replace("'", "''") + "'", _connection);
                                    string sState =
                                        _dbHelper.GetStringbySql(
                                            "SELECT s_Area  FROM dbo.TCstmr_Applicant WHERE s_AppCode='" +
                                            dr["实体ID"].ToString().Trim().Replace("'", "''") + "'", _connection).ToString();
                                    string sStreet =
                                        _dbHelper.GetbySql(
                                            "SELECT s_FirstAddress  FROM dbo.TCstmr_Applicant WHERE s_AppCode='" +
                                            dr["实体ID"].ToString().Trim().Replace("'", "''") + "'", commDB, _connection).ToString();
                                    if (dr["原名"] != null && dr["译名"] != null)
                                    {
                                        string[] s = dr["原名"].ToString().ToUpper().Trim().Split(';');
                                        string[] sN = dr["译名"].ToString().ToUpper().Trim().Split(';');


                                        if (sNativeName.ToUpper().Replace(" ", "").Replace("（", "(").Replace("）", ")").Trim() == s[0].Trim().Replace(" ", "").Replace("（", "(").Replace("）", ")")
                                            || sName.ToUpper().Replace(" ", "").Replace("（", "(").Replace("）", ")").Trim() == sN[0].Trim().Replace(" ", "").Replace("（", "(").Replace("）", ")"))
                                        {
                                            strSql =
                                                "INSERT INTO dbo.TCase_Applicant(s_NativeName,s_Name,s_State,s_City,s_Street,s_AppSerial,n_ApplicantID,n_CaseID) " +
                                                "VALUES('" + sNativeName.Replace("'", "''") + "','" +
                                                sName.Replace("'", "''") + "','" + sState.Replace("'", "''") + "','','" +
                                                sStreet.Replace("'", "''") + "','" +
                                                dr["实体方卷号"].ToString().Trim().Replace("'", "''") + "'," + nApplicantID +
                                                "," + HkNum + ")";
                                            Result = _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                                            if ( Result<= 0)
                                            {
                                                _dbHelper.InsertLog(HkNum, sCaserial, rowid, "国外-实体信息表",
                                                                    "国外-实体信息表-" + rowid,
                                                                    "存在申请人，增加案件申请人失败!实体ID：" + dr["实体ID"].ToString().Trim().Replace("'", "''"), strSql.Replace("'", "''"),
                                                                    commDB, _connection);
                                            }
                                        }
                                        else
                                        {
                                            string msg = "实体ID：" + dr["实体ID"].ToString().Trim().Replace("'", "''");
                                            if ((string) dr["原名"] != sNativeName)
                                            {
                                                msg += "原名:" + dr["原名"] + "--*****原名DB:" + sNativeName;
                                            }
                                            if ((string) dr["译名"] != sName)
                                            {
                                                msg += "\r\n译名:" + dr["译名"] + "--*****译名DB:" + sName;
                                            }

                                            _dbHelper.InsertLog(HkNum, sCaserial, rowid, "国外-实体信息表",
                                                                "国外-实体信息表-申请人-" + rowid,
                                                                "所查到申请人中文名和英文名与原名和译名不同，无法进行更新",
                                                                 msg,
                                                                commDB, _connection);
                                            Result = 0;
                                        }
                                    }
                                }
                            }
                            string strSqlN = "SELECT n_ID FROM dbo.TCase_Applicant  WHERE n_CaseID=" + HkNum +
                                        "  AND n_ApplicantID=" + nAppID;
                            if (_dbHelper.GetbySql(strSqlN, commDB, _connection) <= 0)
                            {
                                Result = _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);  
                                if ( Result  == 0)
                                {
                                    _dbHelper.InsertLog(HkNum, sCaserial, rowid, "国外-实体信息表", "国外-实体信息表-申请人" + rowid, "增加案件申请人失败", strSql.Replace("'", "''"), commDB, _connection);
                                    Result = 0;
                                }
                            }
                        }
                        else
                        {
                            _dbHelper.InsertLog(HkNum, sCaserial, rowid, "国外-实体信息表", "国外-实体信息表-" + rowid, "实体ID为空-申请人：" + dr["实体ID"].ToString().Trim(), "", commDB, _connection);
                            Result = 0;
                        }
                        #endregion

                        #region 申请人及委托人

                        if (dr["实体角色"].ToString().Trim() == "申请人及委托人")
                        {
                            string strSqAl = "select n_ClientID from TCstmr_Client WHERE s_ClientCode='" +
                                             dr["实体ID"].ToString().Trim().Replace("'", "''") + "'";
                            int nClientId = _dbHelper.GetbySql(strSqAl, commDB, _connection);
                            if (nClientId > 0)
                            {
                                strSql = "  update TCase_Base set s_ClientSerial='" + dr["实体方卷号"] +
                                         "',n_ClientID=" + nClientId + " where n_CaseID=" + HkNum;
                               Result= _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                            }
                            else
                            {
                                _dbHelper.InsertLog(HkNum, sCaserial, rowid, "国外-实体信息表", "国外-实体信息表-申请人及委托人" + rowid, "未查到实体ID-申请人及委托人：" + dr["实体ID"].ToString().Trim(), strSql.Replace("'", "''"), commDB, _connection);
                                Result = 0;
                            }
                        } 
                        #endregion
                    }
                    else
                    {
                        _dbHelper.InsertLog(HkNum, sCaserial, rowid, "国外-实体信息表", "国外-实体信息表-" + rowid, "实体ID为空申请人及委托人|申请人", "", commDB, _connection);
                        Result = 0;
                    }

                    #endregion
                }
                else
                {
                    _dbHelper.InsertLog(HkNum, sCaserial, rowid, "国外-实体信息表", "国外-实体信息表-" + rowid, "实体ID为空", strSql.Replace("'", "''"), commDB, _connection);
                    Result = 0;
                } 
               
            }
            return Result;
        }

        #endregion

        #region 更新申请人译名
        public int UpdateApplicant(DataRow dataRow, int row, string commDB, SqlConnection _connection)
        {
            int result = 0;
            string sNo = dataRow["我方文号"].ToString().Trim();
            int numHk = _dbHelper.GetIDbyName(sNo, 2, _connection);
            if (numHk > 0)
            {
                int nApplicantID = _dealingClientData.GetClientandApplicantIDByName(dataRow["申请人编码"].ToString(),
                                                                                    "Applicant", commDB, _connection);
                if (nApplicantID > 0)
                {
                    string strSql = "select  s_Name from TCstmr_Applicant where n_AppID=" + nApplicantID;
                    string sName = _dbHelper.GetStringbySql(strSql, _connection);

                    string strSql2 = "select  s_AppTransLatedName from TCstmr_AppTransLatedName where n_AppID=" +
                                     nApplicantID + " AND s_AppTransLatedName='" + dataRow["申请人译名"].ToString() + "'";
                    string sName2 = _dbHelper.GetStringbySql(strSql2, _connection);

                    if (sName != null && sName.ToUpper().Trim().Replace(" ", "").Replace("（", "(").Replace("）", ")") == dataRow["申请人译名"].ToString().ToUpper().Trim().Replace(" ", "").Replace("（", "(").Replace("）", ")"))
                    {
                        strSql = "select  s_TrustDeedNo from TCstmr_Applicant where n_AppID=" + nApplicantID;
                        string sTrustDeedNo = _dbHelper.GetStringbySql(strSql, _connection);
                        strSql = "UPDATE dbo.TCase_Applicant SET s_Name='" + dataRow["申请人译名"].ToString() +
                                 "',s_TrustDeedNo='" + sTrustDeedNo + "' WHERE n_CaseID=" + numHk +
                                 " AND n_ApplicantID=" + nApplicantID;
                        result = _dbHelper.InsertbySql(strSql, row, commDB, _connection);
                        if (result == 0)
                        {
                            _dbHelper.InsertLog(numHk, sNo, row, "更新申请人译名", "更新申请人译名-" + row,
                                          "更新数据失败，无申请人信息-1：" + dataRow["申请人编码"],
                                          strSql.Replace("'", "''"), commDB,
                                          _connection);
                        }
                    }
                    else if (sName2 != null && sName2.ToUpper().Trim().Replace(" ", "").Replace("（", "(").Replace("）", ")") == dataRow["申请人译名"].ToString().ToUpper().Trim().Replace(" ", "").Replace("（", "(").Replace("）", ")"))
                    {
                        strSql = "select  s_TrustdeedNum from TCstmr_AppTransLatedName where n_AppID=" +
                                 nApplicantID + " AND s_AppTransLatedName='" + dataRow["申请人译名"].ToString() + "'";
                        string sTrustdeedNum = _dbHelper.GetStringbySql(strSql, _connection);
                        strSql = "UPDATE dbo.TCase_Applicant SET s_Name='" + dataRow["申请人译名"].ToString() +
                                 "',s_TrustDeedNo='" + sTrustdeedNum + "' WHERE n_CaseID=" + numHk +
                                 " AND n_ApplicantID=" + nApplicantID;
                        result = _dbHelper.InsertbySql(strSql2, row, commDB, _connection);
                        if(result==0)
                        {
                            _dbHelper.InsertLog(numHk, sNo, row, "更新申请人译名", "更新申请人译名-" + row,
                                          "更新数据失败，无申请人信息-2：" + dataRow["申请人编码"],
                                          strSql.Replace("'", "''") , commDB,
                                          _connection); 
                        }
                    }
                    else
                    {
                        _dbHelper.InsertLog(numHk, sNo, row, "更新申请人译名", "更新申请人译名-" + row,
                                            dataRow["申请人编码"].ToString() + "用户下，不存在译名为：" +
                                            dataRow["申请人译名"].ToString(),
                                            strSql.Replace("'", "''") + "&&&" + strSql2.Replace("'", "''"), commDB,
                                            _connection);
                    }
                }
                else
                {
                    _dbHelper.InsertLog(numHk, sNo, row, "更新申请人译名", "更新申请人译名-" + row,
                                        "未找到申请人为:" + dataRow["申请人编码"].ToString(), "", commDB, _connection);
                }
            }
            else
            {
                _dbHelper.InsertLog(numHk, sNo, row, "更新申请人译名", "更新申请人译名-" + row, "未找到“我方卷号”为:" + sNo, "", commDB,
                                    _connection);
            }

            return result;
        }
        #endregion
    }
}
