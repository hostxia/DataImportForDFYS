using System;
using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport.Comm
{
    internal class dealingCaseDemand
    {
        private readonly DBHelper _dbHelper = new DBHelper();

        public int InsertDemand(string nCaseID, int rowid, string commDB, SqlConnection _connection)
        { 
            int result = 0;
            string strSql = " select n_ClientID from TCase_Clients where n_CaseID=" + nCaseID + "";
            DataTable table = _dbHelper.GetDataTablebySql(strSql, _connection);
            for (int k = 0; k < table.Rows.Count; k++)
            {
                int n_ClientID = int.Parse(table.Rows[k]["n_ClientID"].ToString());
                result = InDemand("相关客户", n_ClientID, nCaseID, commDB, _connection);
            }
            return result;
        }

        private int InDemand(string type, int nClientID, string nCaseID, string commDB, SqlConnection _connection)
        {
            int result = 0;
            string sModuleType = "Client";
            string strSql = "select n_ID from T_Demand where n_ApplicantID IS NULL AND s_SysDemand IS NULL";

            if (type.Equals("相关客户"))
            {
                strSql += "  and  n_ClientID=" + nClientID;
                sModuleType = "RelatedClient";
            }
            DataTable newtable = _dbHelper.GetDataTablebySql(strSql, _connection);
            for (int i = 0; i < newtable.Rows.Count; i++)
            {
                result = AddCase(int.Parse(newtable.Rows[i]["n_ID"].ToString()), nCaseID, sModuleType, commDB, _connection);
            }
            return result;
        }

        private int AddCase(int nID, string nCaseID, string moduleType, string commDB, SqlConnection _connection)
        {
            //根据系统要求代码查询主题和描述
            string strl =
                "select n_ID,s_IPtype,s_Title,s_Description,n_DemandType,n_SysDemandID,n_CodeDemandID from T_Demand WHERE n_ID=" +
                nID;

            DataTable newTable = _dbHelper.GetDataTablebySql(strl, _connection);

            for (int k = 0; k < newTable.Rows.Count; k++)
            {
                string n_ID = newTable.Rows[k]["n_ID"].ToString();
                string n_DemandType = newTable.Rows[k]["n_DemandType"].ToString();
                string s_IPType = newTable.Rows[k]["s_IPtype"].ToString();
                string title = newTable.Rows[k]["s_Title"].ToString();
                string description = newTable.Rows[k]["s_Description"].ToString();
                string n_SysDemandID = newTable.Rows[k]["n_SysDemandID"].ToString();

                string Sql =
                    "INSERT INTO dbo.T_Demand(s_sourcetype1,s_ModuleType,s_Title,s_Description,s_Creator,s_Editor,s_IPType,n_DemandType,dt_EditDate,dt_CreateDate,n_SysDemandID,n_CodeDemandID,s_SourceModuleType,n_SourceID,n_CaseID)" +
                    "VALUES  ('7.相关客户案件要求','Case','" + title + "','" + description +
                    "','administrator','administrator','" + s_IPType + "','" + n_DemandType + "','" + DateTime.Now +
                    "','" + DateTime.Now + "'," + n_SysDemandID + "," + n_SysDemandID + ",'" + moduleType + "'," + n_ID +
                    ",'" + nCaseID + "')";

                //查询是否存在此案件要求要求
                string strSql =
                    "select n_ID,s_SourceModuleType from T_Demand where s_ModuleType='Case'  and n_CaseID=" + nCaseID +
                    " and  n_SysDemandID=" + n_SysDemandID;
                DataTable Table = _dbHelper.GetDataTablebySql(strSql, _connection);

                if (Table.Rows.Count <= 0)
                {
                    return _dbHelper.InsertbySql(Sql, 0, commDB, _connection);
                }
                else
                {
                    string Type = Table.Rows[0]["s_SourceModuleType"].ToString();
                   if (moduleType.Equals("RelatedClient") && (Type.Equals("Applicant") || Type.Equals("Client")))
                    {
                        strSql = "update T_Demand set s_sourcetype1='7.相关客户案件要求',dt_EditDate='" + DateTime.Now + "',s_SourceModuleType='" +
                                 moduleType + "' where s_ModuleType='Case'  and n_CaseID=" + nCaseID +
                                 " and  n_SysDemandID=" + n_SysDemandID;
                    }
                   return _dbHelper.InsertbySql(strSql, 0, commDB, _connection);
                }
            }
            return 0;
        }
    }
}
