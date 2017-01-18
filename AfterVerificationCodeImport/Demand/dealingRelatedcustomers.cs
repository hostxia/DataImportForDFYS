using System.Data;
using System.Data.SqlClient;
using AfterVerificationCodeImport.Four;

namespace AfterVerificationCodeImport.Demand
{
    class dealingRelatedcustomers
    {
        readonly DBHelper _dbHelper = new DBHelper();

        //相关客户
        public int UpdateRelatedcustomers(DataRow row, int rowid, string commDB, SqlConnection _connection)
        {
            int result = 0;
            string sNo = row["客户代码"].ToString();
            string relatedcustomers = row["相关客户代码"].ToString();

            if (!string.IsNullOrEmpty(sNo) && !string.IsNullOrEmpty(relatedcustomers) && relatedcustomers.Trim() != "0")
            {
                //1.查询出相关客户代码系统ID 
                string strSql = "select  n_ClientID from TCstmr_Client where s_ClientCode='" + sNo + "'";
                int Num = _dbHelper.GetbySql(strSql, commDB, _connection);
                if (Num > 0)
                {
                    //1.根据客户代码查询出所有客户案件
                    strSql = "select n_CaseID from TCase_Base  where n_ClientID in (select  n_ClientID from TCstmr_Client where s_ClientCode='" + sNo + "')";
                    DataTable table = _dbHelper.GetDataTablebySql(strSql,_connection);
                    strSql = "select n_CaseID from TCase_Applicant  where n_ApplicantID in(select n_AppID from TCstmr_Applicant  where s_AppCode='" + sNo + "')";
                    DataTable newtable = _dbHelper.GetDataTablebySql(strSql, _connection);
                    table.Merge(newtable);
                    var _dealingCasePantent = new dealingCasePantent();
                    for (int k = 0; k < table.Rows.Count; k++)
                    {
                        _dealingCasePantent.InsertTCaseClients(relatedcustomers, int.Parse(table.Rows[k]["n_CaseID"].ToString()), rowid, "相关客户-集团客户代码", commDB, _connection);
                    }
                    result = 1;
                }
                else
                {
                    _dbHelper.InsertLog(0, "", rowid, "相关客户-集团客户代码", "相关客户-集团客户代码-" + rowid, "为查到此客户信息：" + sNo, "", commDB, _connection);
                }
            }
            return result;
        }

        //集团客户代码-相关客户要求增加到客户和申请人内
        public int InsertClientDemnd(DataRow row, int rowid, string commDB, SqlConnection _connection)
        {
            string strSql = "select  n_ClientID from TCstmr_Client where s_ClientCode='" + row["客户代码"].ToString() + "'";
            int nClientIDA = _dbHelper.GetbySql(strSql, commDB, _connection);
            strSql = "select  n_AppID from TCstmr_Applicant where s_AppCode='" + row["申请人代码"].ToString() + "'";
            int nAppID = _dbHelper.GetbySql(strSql, commDB, _connection);

            strSql = "select  n_ClientID from TCstmr_Client where s_ClientCode='" + row["相关客户代码"].ToString() + "'";
            int nClientID = _dbHelper.GetbySql(strSql, commDB, _connection);

            string strl = "select n_ID,s_IPtype,s_Title,s_Description,s_Creator,n_DemandType,n_SysDemandID,n_CodeDemandID,s_ModuleType,s_sysDemand from T_Demand WHERE n_ApplicantID IS NULL AND s_SysDemand IS NULL and n_ClientID=" + nClientID;
            var newTable = _dbHelper.GetDataTablebySql(strl, _connection);

            for (int k = 0; k < newTable.Rows.Count; k++)
            {
                string nDemandType = newTable.Rows[k]["n_DemandType"].ToString();
                string sIPType = newTable.Rows[k]["s_IPtype"].ToString();
                string title = newTable.Rows[k]["s_Title"].ToString();
                string description = newTable.Rows[k]["s_Description"].ToString();
                string sCreator = newTable.Rows[k]["s_Creator"].ToString();

                strSql = "select n_ID from T_Demand WHERE n_ClientID=" + nClientIDA + " AND s_Title='" + title + "' AND s_Description='" + description + "' AND n_DemandType='" + nDemandType + "' and s_IPType='" + sIPType + "'";
                if (_dbHelper.GetbySql(strSql, commDB, _connection) <= 0)//客户
                {
                    if (nClientIDA > 0)
                    {
                        strSql =
                            "INSERT INTO dbo.T_Demand(s_sourcetype1,s_ModuleType,s_Title,s_Description,s_Creator,s_IPType,n_DemandType,n_ClientID)" +
                            "VALUES  ('相关客户-要求-客户','Client','" + title + "','" + description + "','" + sCreator + "','" +
                            sIPType + "'," + nDemandType + "," + nClientIDA + ")";
                        _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                    }
                    else
                    {
                        _dbHelper.InsertLog(0, "", rowid, "相关客户-集团客户代码", "相关客户-要求-客户-" + rowid, "为查到此客户信息：" + row["客户代码"], "", commDB, _connection);
                    }
                }
                strSql = "select n_ID from T_Demand WHERE n_ApplicantID=" + nAppID + " AND s_Title='" + title + "' AND s_Description='" + description + "' AND n_DemandType='" + nDemandType + "' and s_IPType='" + sIPType + "'";
                if (_dbHelper.GetbySql(strSql, commDB, _connection) <= 0)//申请人
                {
                    if (nAppID > 0)
                    {
                        strSql =
                            "INSERT INTO dbo.T_Demand(s_sourcetype1,s_ModuleType,s_Title,s_Description,s_Creator,s_IPType,n_DemandType,n_ApplicantID)" +
                            "VALUES  ('相关客户-要求-申请人','Applicant','" + title + "','" + description + "','" + sCreator +
                            "','" + sIPType + "'," + nDemandType + "," + nAppID + ")";
                        _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                    }
                    else
                    {
                        _dbHelper.InsertLog(0, "", rowid, "相关客户-集团客户代码", "相关客户-要求-申请人-" + rowid, "为查到此申请人信息：" + row["客户代码"], "", commDB, _connection);
                    }
                }
            }
            return 1;
        } 
    }
}
