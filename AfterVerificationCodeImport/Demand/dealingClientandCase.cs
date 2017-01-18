using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace AfterVerificationCodeImport.Demand
{
    class dealingClientandCase
    {
        readonly DBHelper _dbHelper = new DBHelper();

        #region  要求配置
        public int InsertDemand(DataRow row, int rowid, string commDB, SqlConnection _connection)
        {
            var Result = 0;
            string sNo = row["客户ID"].ToString();
            string relatedcustomers = row["申请人ID"].ToString();
            string demand = row["要求ID"].ToString();
            string strSql = "select top 1 n_ClientID from TCstmr_Client where s_ClientCode='" + sNo + "'";
            int nClientID = _dbHelper.GetbySql(strSql, commDB, _connection);
            strSql = " select n_AppID from TCstmr_Applicant  where s_AppCode='" + relatedcustomers + "'";
            int nAppID = _dbHelper.GetbySql(strSql, commDB, _connection);

            if (nClientID > 0 && nAppID>0) //客户申请人
            {
                InsertCodeDemands(demand, "ClientAppliant", nClientID, nAppID, commDB, _connection);
            }
            else
            {
                if (nClientID > 0) //增加客户和同客户申请人
                {
                    InsertCodeDemands(demand, "Client", nClientID, 0, commDB, _connection);

                    //查询客户ID是否也存在申请人表内
                    strSql = " select n_AppID from TCstmr_Applicant  where s_AppCode='" + sNo + "'";
                    nAppID = _dbHelper.GetbySql(strSql, commDB, _connection);
                    if (nAppID > 0)
                    {
                        InsertCodeDemands(demand, "Applicant", 0, nAppID, commDB, _connection);
                    }
                }
                if (nAppID > 0)
                {
                    //查询客户ID是否也存在申请人表内
                    if (nAppID > 0)
                    {
                        InsertCodeDemands(demand, "Applicant", 0, nAppID, commDB, _connection);
                    }
                }
                Result= 1;
            }
            return Result;
        }

        //客户、申请人要求
        private void InsertCodeDemands(string demand, string moduleType, int nClientID, int nApplicantID, string commDB, SqlConnection _connection)
        {
                //根据系统要求代码查询主题和描述
                DataTable newTable = _dbHelper.GetDataTablebySql("select n_ID,s_IPtype,s_Title,s_Description,n_DemandType from TCode_Demand WHERE s_sysdemand='" + demand + "'", _connection);
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
                InsertDemand(demand, moduleType, nClientID, nApplicantID, n_ID, title, description, s_IPType, n_DemandType, commDB, _connection); 
        }
        private void InsertDemand(string demand, string moduleType, int nClientID, int nApplicantID, string n_ID, string title, string description, string s_IPType, string n_DemandType, string commDB, SqlConnection _connection)
        { 
            DateTime time = DateTime.Now;
            //插入客户要求并且重复插入申请人要求
            string SelectSql = "select n_ID from T_Demand where  n_SysDemandID='" + n_ID + "'";//s_ModuleType='" + moduleType + "' and
            string Sql = "INSERT INTO dbo.T_Demand(s_sourcetype1,s_ModuleType,s_Title,s_Description,s_Creator,s_Editor,s_IPType,s_SysDemand,n_DemandType,dt_EditDate,dt_CreateDate,n_SysDemandID,n_CodeDemandID";
            string Sql2 = "VALUES  ('客户要求代码导入--客户/申请人','" + moduleType + "','" + title + "','" + description + "','administrator','administrator','" + s_IPType + "','" + demand + "','" + n_DemandType + "','" + time + "','" + time + "'," + n_ID + "," + n_ID;

            if (moduleType.Equals("Applicant"))//申请人
            {
                Sql += ",n_ApplicantID";
                Sql2 += "," + nApplicantID;
                SelectSql += " and n_ApplicantID='" + nApplicantID + "'";
            }
            else if (moduleType.Equals("Client"))//客户
            {
                Sql += ",n_ClientID";
                Sql2 += "," + nClientID;
                SelectSql += " and n_ClientID='" + nClientID + "'";
            }
            else//客户-申请人ClientApplicant
            {
                Sql += ",n_ApplicantID,n_ClientID";
                Sql2 += "," + nApplicantID + "," + nClientID;
                SelectSql += " and n_ClientID='" + nClientID + "' and n_ApplicantID='" + nApplicantID + "'";
            }
            Sql += ")";
            Sql2 += ")";
            int Num = _dbHelper.GetbySql(SelectSql, commDB, _connection);
            if (Num > 0)
            {
                string strSql = "update T_Demand set  s_ModuleType='" + moduleType + "' where n_ID=" + Num;
                _dbHelper.InsertbySql(strSql, 0, commDB, _connection);
            }
            else
            {
                _dbHelper.InsertbySql(Sql + Sql2, 0, commDB, _connection);
                _dbHelper.GetbySql(SelectSql, commDB, _connection);
            }
        }
        #endregion 

        #region 增加案件要求
        //案件要求Copy
        public void InsertCaseDemand(DataRow row, string commDB, SqlConnection _connection)
        {
            string sNo = row["客户ID"].ToString();
            string relatedcustomers = row["申请人ID"].ToString();
            string demand = row["要求ID"].ToString();

            string strSql = "select  n_ClientID from TCstmr_Client where s_ClientCode='" + sNo + "'";
            int nClientID = _dbHelper.GetbySql(strSql, commDB, _connection);
            strSql = "select  n_ClientID from TCstmr_Client where s_ClientCode='" + relatedcustomers + "'";
            int nClientID2 = _dbHelper.GetbySql(strSql, commDB, _connection);

            strSql = "select  n_AppID from TCstmr_Applicant where s_AppCode='" + relatedcustomers + "'";
            int nAppID = _dbHelper.GetbySql(strSql, commDB, _connection);
            int nDemand = _dbHelper.GetbySql("select top 1 n_ID from TCode_Demand WHERE s_sysdemand='" + demand + "'", commDB, _connection);

            //相关客户TCase_Clients
            strSql = "select n_CaseID,n_ClientID from TCase_Clients where n_ClientID in (" + nClientID + "," + nClientID2 + ")";
            DataTable table = _dbHelper.GetDataTablebySql(strSql,_connection);
            for (int k = 0; k < table.Rows.Count; k++)
            {
                if (SelectClient(nClientID, nClientID2, nAppID, int.Parse(table.Rows[k]["n_CaseID"].ToString()), _connection))
                {
                    nDemand = 0;
                }
                InCase(table.Rows[k]["n_CaseID"].ToString(), "相关客户", nDemand, commDB, _connection); 
            }

            //根据申请人查询增加要求
            if (nAppID > 0)
            {
                strSql = "select  n_CaseID from TCase_Applicant  where n_ApplicantID=" + nAppID;
                table = _dbHelper.GetDataTablebySql(strSql,_connection);

                for (int k = 0; k < table.Rows.Count; k++)
                {
                    if (SelectClient(nClientID, nClientID2, nAppID, int.Parse(table.Rows[k]["n_CaseID"].ToString()), _connection))
                    {
                        nDemand = 0;
                    }
                    InCase(table.Rows[k]["n_CaseID"].ToString(), "申请人", nDemand, commDB, _connection); 
                }
            }

            //根据客户查询增加要求
            if (nClientID > 0)
            {
                strSql = "select n_CaseID from TCase_Base  where n_ClientID in (select  n_ClientID from TCstmr_Client where n_ClientID='" + nClientID + "')";
                table = _dbHelper.GetDataTablebySql(strSql,_connection);
                for (int k = 0; k < table.Rows.Count; k++)
                {
                    if (SelectClient(nClientID, nClientID2, nAppID, int.Parse(table.Rows[k]["n_CaseID"].ToString()), _connection))
                    {
                        nDemand = 0;
                    }
                    InCase(table.Rows[k]["n_CaseID"].ToString(), "客户", nDemand, commDB, _connection); 
                }
            }
        }
        
        //如果客户和申请人都存在，则导入要求，不存在则不导入
        private bool SelectClient(int nClientID, int nClientID2, int nApplicantID, int nCaseID, SqlConnection _connection)
        {
            string strSql = "select n_CaseID from TCase_Clients where n_ClientID in (" + +nClientID + "," + nClientID2 + ")";
            DataTable table = _dbHelper.GetDataTablebySql(strSql, _connection);
            strSql = "select n_CaseID from TCase_Base  where n_ClientID in (select  n_ClientID from TCstmr_Client where n_ClientID='" + nClientID + "')";
            DataTable table3 = _dbHelper.GetDataTablebySql(strSql, _connection);
            table.Merge(table3);
            int row = table.Select("n_CaseID ='" + nCaseID + "' ").ToList().Count();


            strSql = "select  n_CaseID from TCase_Applicant  where n_ApplicantID=" + nApplicantID + " and n_CaseID=" + nCaseID;
            DataTable table2 = _dbHelper.GetDataTablebySql(strSql, _connection);
            if (nClientID2.Equals(0) || (row > 0 && table2.Rows.Count > 0))
            {
                return true;
            }
            return false;
        }

        private void InCase(string nCaseID, string type, int nDemand, string commDB, SqlConnection _connection)
        {
            try
            {

                //增加相关客户要求  
                string strSql = " select n_ClientID from TCase_Clients where n_CaseID=" + nCaseID + "";
                DataTable table = _dbHelper.GetDataTablebySql(strSql, _connection);
                for (int k = 0; k < table.Rows.Count; k++)
                {
                    if (!string.IsNullOrEmpty(table.Rows[k]["n_ClientID"].ToString()))
                    {
                        int n_ClientID = int.Parse(table.Rows[k]["n_ClientID"].ToString());
                        InDemand("相关客户", n_ClientID, nCaseID, nDemand, commDB, _connection);
                    }
                }
                strSql = "select n_ClientID from TCase_Base where n_CaseID=" + nCaseID;
                table = _dbHelper.GetDataTablebySql(strSql, _connection);
                for (int k = 0; k < table.Rows.Count; k++)
                {
                    if (!string.IsNullOrEmpty(table.Rows[k]["n_ClientID"].ToString()))
                    {
                        int n_ClientID = int.Parse(table.Rows[k]["n_ClientID"].ToString());
                        InDemand("客户", n_ClientID, nCaseID, nDemand, commDB, _connection);
                    }
                }
                strSql = "select  n_ApplicantID from TCase_Applicant  where n_CaseID=" + nCaseID;
                table = _dbHelper.GetDataTablebySql(strSql, _connection);
                for (int k = 0; k < table.Rows.Count; k++)
                {
                    if (!string.IsNullOrEmpty(table.Rows[k]["n_ClientID"].ToString()))
                    {
                        int n_ClientID = int.Parse(table.Rows[k]["n_ApplicantID"].ToString());
                        InDemand("申请人", n_ClientID, nCaseID, nDemand, commDB, _connection);
                    }
                }
            }
            catch (Exception exception)
            {
                string d = exception.Message;
            }
        }

        private void InDemand(string type, int nClientID, string nCaseID, int nDemand, string commDB, SqlConnection _connection)
        {
            try
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
            DataTable newtable = _dbHelper.GetDataTablebySql(strSql, _connection);
            for (int i = 0; i < newtable.Rows.Count; i++)
            {
                AddCase(int.Parse(newtable.Rows[i]["n_ID"].ToString()), nCaseID, sModuleType, commDB, _connection);
            }
            }
            catch (Exception exception)
            {
                string d = exception.Message;
            }
        }

        private void AddCase(int nID, string nCaseID, string moduleType, string commDB, SqlConnection _connection)
        {
            try
            {
                //根据系统要求代码查询主题和描述
                string strl =
                    "select n_ID,s_IPtype,s_Title,s_Description,s_Creator,n_DemandType,n_SysDemandID,n_CodeDemandID,s_ModuleType,s_sysDemand from T_Demand WHERE n_ID=" +
                    nID;

                DataTable newTable = _dbHelper.GetDataTablebySql(strl, _connection);

                string s_IPType = string.Empty;
                string title = string.Empty;
                string description = string.Empty;
                string sCreator = string.Empty;
                string n_DemandType = string.Empty;

                string n_ID = string.Empty;
                string s_sysDemand = string.Empty;
                string n_SysDemandID = "0";
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
                        if (!string.IsNullOrEmpty(newTable.Rows[k]["n_SysDemandID"].ToString()))
                        {
                            n_SysDemandID = newTable.Rows[k]["n_SysDemandID"].ToString();
                        }

                        string Sql =
                            "INSERT INTO dbo.T_Demand(s_sourcetype1,s_ModuleType,s_Title,s_Description,s_Creator,s_Editor,s_IPType,s_SysDemand,n_DemandType,dt_EditDate,dt_CreateDate,s_SourceModuleType,n_SourceID,n_CaseID";
                        string sqlValue = "VALUES  ('客户要求代码-案件','Case','" + title.Replace("'", "''") + "','" + description.Replace("'", "''") + "','" +
                                          sCreator + "','" + sCreator + "','" + s_IPType + "','" + s_sysDemand + "','" +
                                          n_DemandType + "','" + DateTime.Now + "','" + DateTime.Now + "','" +
                                          moduleType + "'," + n_ID + ",'" + nCaseID + "'";
                        if (!string.IsNullOrEmpty(n_SysDemandID) && !n_SysDemandID.Equals("0"))
                        {
                            Sql += ",n_CodeDemandID,n_SysDemandID";
                            sqlValue += "," + n_SysDemandID + "," + n_SysDemandID;
                        }
                        Sql += ")";
                        sqlValue += ")";

                        //查询是否存在此案件要求要求
                        if (!string.IsNullOrEmpty(n_SysDemandID) && !n_SysDemandID.Equals("0"))
                        {
                            string strSql =
                                "select n_ID,s_SourceModuleType from T_Demand where s_ModuleType='Case'  and n_CaseID=" +
                                nCaseID + " and  n_SysDemandID=" + n_SysDemandID;
                            DataTable Table = _dbHelper.GetDataTablebySql(strSql, _connection);

                            if (Table.Rows.Count <= 0)
                            {
                                _dbHelper.InsertbySql(Sql + sqlValue, 0, commDB, _connection);
                            }
                            else
                            {
                                string Type = Table.Rows[0]["s_SourceModuleType"].ToString();
                                if (moduleType.Equals("Applicant") && Type.Equals("Client"))
                                {
                                    strSql = "update T_Demand set dt_EditDate='" + DateTime.Now +
                                             "', s_SourceModuleType='" +
                                             moduleType + "' where s_ModuleType='Case'  and n_CaseID=" + nCaseID +
                                             " and  n_SysDemandID=" + n_SysDemandID;
                                }
                                else if (moduleType.Equals("RelatedClient") &&
                                         (Type.Equals("Applicant") || Type.Equals("Client")))
                                {
                                    strSql = "update T_Demand set dt_EditDate='" + DateTime.Now +
                                             "',s_SourceModuleType='" + moduleType +
                                             "' where s_ModuleType='Case'  and n_CaseID=" + nCaseID +
                                             " and  n_SysDemandID=" + n_SysDemandID;
                                }
                                _dbHelper.InsertbySql(strSql, 0, commDB, _connection);
                            }
                        }
                        else
                        {
                            string strSql =
                              "select n_ID,s_SourceModuleType from T_Demand where s_ModuleType='Case'  and n_CaseID=" +
                              nCaseID + " and  s_Title='" + title.Replace("'", "''") + "' and  s_Description='" + description.Replace("'", "''") + "' and n_DemandType='" + n_DemandType + "' and n_SourceID='" + n_ID + "'";
                            DataTable Table = _dbHelper.GetDataTablebySql(strSql, _connection);
                             
                            if (Table.Rows.Count <= 0)
                            {
                                _dbHelper.InsertbySql(Sql + sqlValue, 0, commDB, _connection);
                            }
                            else
                            {
                                string Type = Table.Rows[0]["s_SourceModuleType"].ToString();
                                if (moduleType.Equals("Applicant") && Type.Equals("Client"))
                                {
                                    strSql = "update T_Demand set dt_EditDate='" + DateTime.Now +
                                             "', s_SourceModuleType='" +
                                             moduleType + "' where s_ModuleType='Case'  and n_CaseID=" + nCaseID +
                                             " and  n_SysDemandID=" + n_SysDemandID;
                                }
                                else if (moduleType.Equals("RelatedClient") &&
                                         (Type.Equals("Applicant") || Type.Equals("Client")))
                                {
                                    strSql = "update T_Demand set dt_EditDate='" + DateTime.Now +
                                             "',s_SourceModuleType='" + moduleType +
                                             "' where s_ModuleType='Case'  and n_CaseID=" + nCaseID +
                                             " and  n_SysDemandID=" + n_SysDemandID;
                                }
                                _dbHelper.InsertbySql(strSql, 0, commDB, _connection);
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                string d = exception.Message;
            }
        }

        #endregion

    }
}
