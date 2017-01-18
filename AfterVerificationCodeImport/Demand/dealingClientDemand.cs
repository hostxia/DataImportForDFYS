using System;
using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport.Demand
{
    class dealingClientDemand
    {
        readonly DBHelper _dbHelper = new DBHelper();

        public int InsertDemandClient(DataRow dr, int rowid, string commDB, SqlConnection _connection)
        {
            int nDemandType = InsertDemandType(dr["标题"].ToString(), commDB, _connection);//分类ID
            if (dr["是否联合要求"] != null && dr["客户编号"].ToString().ToUpper().Equals("Y"))
            {
                string sClientCode = dr["客户编号"].ToString() + "-" + dr["申请人编号"].ToString();
                int nClientID =
                      _dbHelper.GetbySql(
                          "SELECT n_ClientID FROM  dbo.TCstmr_Client   WHERE s_ClientCode='" + dr["客户编号"].ToString() + "'",
                          commDB, _connection);
                int nAppID =
                      _dbHelper.GetbySql(
                          "SELECT n_AppID FROM  dbo.TCstmr_Applicant  WHERE s_AppCode='" + dr["申请人编号"].ToString() + "'",
                          commDB, _connection);
                if (nClientID > 0 && nAppID > 0)
                {
                    InsertUserDemandClient(nDemandType, dr["收到方式"].ToString(), dr["指示人"].ToString(), dr["标题"].ToString(),
                                           dr["描述"].ToString(), dr["收到日"].ToString(), "客户申请人", sClientCode, rowid,
                                           nClientID, nAppID, commDB, _connection);
                    return 1;
                }
                else
                {
                    _dbHelper.InsertLog(0, "", rowid, "客户要求配置", "客户要求配置-" + rowid, "是否联合要求：Y,缺少客户" + dr["客户编号"].ToString() + "-" + nClientID + "或者申请人" + dr["申请人编号"].ToString()+"-"+ nAppID, "",
                                               commDB, _connection);
                    return 0;
                }
            }
            else
            {
                string sClientCode = "0";
                if (dr["客户编号"] != null && !string.IsNullOrEmpty(dr["客户编号"].ToString()))
                {
                    sClientCode = dr["客户编号"].ToString();
                    int nClientID =
                        _dbHelper.GetbySql(
                            "SELECT n_ClientID FROM  dbo.TCstmr_Client   WHERE s_ClientCode='" + sClientCode + "'",
                            commDB, _connection);
                    if (nClientID > 0)
                    {
                        InsertUserDemandClient(nDemandType, dr["收到方式"].ToString(), dr["指示人"].ToString(), dr["标题"].ToString(),
                                               dr["描述"].ToString(), dr["收到日"].ToString(), "客户", sClientCode, rowid,
                                               nClientID, 0, commDB, _connection);
                    }
                    else
                    {
                        _dbHelper.InsertLog(0, "", rowid, "客户要求配置", "客户要求配置-" + rowid, "不存在客户为：" + sClientCode, "",
                                             commDB, _connection);
                    }
                }
                if (dr["申请人编号"] != null && !string.IsNullOrEmpty(dr["申请人编号"].ToString()))
                {
                    sClientCode = dr["申请人编号"].ToString();
                    int nAppID =
                        _dbHelper.GetbySql(
                            "SELECT n_AppID FROM  dbo.TCstmr_Applicant  WHERE s_AppCode='" + sClientCode + "'",
                            commDB, _connection);
                    if (nAppID > 0)
                    {
                        InsertUserDemandClient(nDemandType,dr["收到方式"].ToString(), dr["指示人"].ToString(), dr["标题"].ToString(),
                                               dr["描述"].ToString(), dr["收到日"].ToString(), "申请人", sClientCode, rowid,
                                               0, nAppID, commDB, _connection);
                    }
                    else
                    {
                        _dbHelper.InsertLog(0, "", rowid, "客户要求配置", "客户要求配置-" + rowid, "不存在申请人为：" + sClientCode, "",
                                             commDB, _connection);
                    }
                }
                return 1;
            }
        }
        private int InsertDemandType(string sCodeDemandType, string commDB, SqlConnection _connection)
        {
            string strSql = "select n_ID from TFCode_DemandType where s_Name='" + sCodeDemandType + "'";
            int num = _dbHelper.GetbySql(strSql, commDB, _connection);

            if (num <= 0 )
            {
                string InsertSql = "INSERT INTO dbo.TFCode_DemandType( s_Name)"
                                   + "VALUES  ( '" + sCodeDemandType + "')";
                if (_dbHelper.InsertbySql(InsertSql, 0, commDB, _connection) > 0)
                {
                    num = _dbHelper.GetbySql(strSql, commDB, _connection);
                }
            } 
            return num;
        }


        private void InsertUserDemandClient(int nDemandType,string sReceiptMethod, string sAssignor, string sTitle,
                                     string sDescription, string dtReceiptDate, string type, string sClientCode, int rowid, int nClientCodeID, int nApplicantID, string commDB, SqlConnection _connection)
        {
            string selectSql = "SELECT n_ID FROM dbo.T_Demand WHERE s_Title='" + sTitle.Replace("'", "''") + "' and s_Description='" + sDescription.Replace("'", "''") + "' and n_DemandType=" + nDemandType;
            string strSql =
                "INSERT INTO dbo.T_Demand(s_sourcetype1,n_DemandType,dt_CreateDate,dt_EditDate,s_ReceiptMethod,s_Assignor,s_Title,s_Description";
            string strSql1 = " VALUES  ('客户要求-" + type + "-" + rowid + "'," + nDemandType + ",'" + DateTime.Now + "','" + DateTime.Now + "','" +
                             sReceiptMethod.Replace("'", "''") + "','" + sAssignor.Replace("'", "''") + "','" +
                             sTitle.Replace("'", "''") + "','" + sDescription.Replace("'", "''") + "'";
            if (!string.IsNullOrEmpty(dtReceiptDate))
            {
                strSql += ",dt_ReceiptDate";
                strSql1 += ",'" + dtReceiptDate + "'";
                selectSql += " and dt_ReceiptDate='" + dtReceiptDate +"'";
            }
            if (type.Equals("申请人"))
            {
                strSql += ",s_ModuleType,n_ApplicantID";
                strSql1 += ",'Applicant'," + nApplicantID;
                selectSql += " and s_ModuleType='Applicant' and Applicant=" + nApplicantID; 
            }
            else if (type.Equals("客户"))
            {
                strSql += ",s_ModuleType,n_ClientID ";
                strSql1 += ",'Client'," + nClientCodeID;
                selectSql += " and s_ModuleType='Client' and n_ClientID=" + nClientCodeID; 
            }
            else if (type.Equals("客户申请人"))
            {
                strSql += ",s_ModuleType,n_ClientID ,n_ApplicantID";
                strSql1 += ",'ClientApplicant'," + nClientCodeID + "," + nApplicantID;
                selectSql += " and s_ModuleType='ClientApplicant' and Applicant=" + nApplicantID + " n_ClientID=" + nClientCodeID; 
            }

            strSql += ")";
            strSql1 += ")";
            string sql = strSql + strSql1;
            if (_dbHelper.GetbySql(selectSql, commDB, _connection) <= 0)
            {
                if (_dbHelper.InsertbySql(strSql + strSql1, rowid, commDB, _connection) <= 0)
                {
                    _dbHelper.InsertLog(0, "", rowid, "客户要求配置", "客户要求配置-" + type + "-" + rowid, "插入要求报错：" + sClientCode,
                                        (sql + strSql1).Replace("'", "''"), commDB, _connection);
                }
            }
        }
    }
}
