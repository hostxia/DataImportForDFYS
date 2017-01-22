using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AfterVerificationCodeImport
{
    class DBHelper
    {
        //执行数据库查询
        public int GetbySql(string strSql, string commDB, SqlConnection conn)
        {
            int retName = 0;
            var sqlColumns = new List<string>();
            strSql = strSql.Replace("\r", "").Replace("\n", "");
            try
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @strSql;
                    cmd.Parameters.Add(new SqlParameter("@name", commDB));
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

        /// 更加SQL执行增、删、改,带有行号 
        public int InsertbySql(string strSql, int iRow, string tableName, SqlConnection conn)
        {
            int retName = 0;
            try
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @strSql;
                    cmd.Parameters.Add(new SqlParameter("@name", tableName));
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

        //增加日志数据
        public int InsertLog(int _caseID, string _caseSerial, int _RowID, string _tabName, string _excelName, string _msg, string _strSql, string tableName, SqlConnection conn)
        {
            return InsertbySql("INSERT INTO [dbo].[Table_Log]([n_CaseID],[s_CaseSerial],[n_RowID],tableName,ExcelName,msg,strsql) " +
                                      "VALUES(" + _caseID + ",'" + _caseSerial.Replace("'", "''") + "'," + _RowID + ",'" + _tabName + "','" + _excelName + "-" + _RowID + "','" + _msg.Replace("'", "''") + "','" + _strSql.Replace("'", "''") + "')", _RowID, tableName, conn);
        }

        /// 查询我方卷号  国家
        public int GetIDbyName(string rowName, int type, SqlConnection conn)
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
                        cmd.CommandText = @"select n_CaseID from TCase_Base where s_CaseSerial='" + rowName + "'";
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
                #region 查找申请人信息

                try
                {
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        string[] s = rowName.Split(';');
                        cmd.CommandText = @"SELECT n_AppID FROM TCstmr_Applicant WHERE s_AppCode='" + s[0] + "'";
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
                        cmd.CommandText =
                            @"SELECT n_CaseID FROM dbo.TPCase_Patent WHERE n_CaseID IN (SELECT n_CaseID from TCase_Base where s_CaseSerial='" +
                            rowName + "')";
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
                        cmd.CommandText = @"select n_AgencyID  from TCstmr_CoopAgency WHERE s_Code='" + rowName + "'";
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
                        cmd.CommandText = @"SELECT TOP 1 n_ID FROM TCstmr_AgencyAddress WHERE n_AgencyID='" + rowName +
                                          "'  ORDER BY n_ID ASC";
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
                        cmd.CommandText = @"SELECT n_CaseID FROM dbo.TCase_Base WHERE  s_AppNo='" + rowName +
                                          "'  ORDER BY n_CaseID ASC";
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

        /// 根据传入sql查询固定值 
        public DataTable GetDataTablebySql(string strSql, SqlConnection conn)
        {
            var table = new DataTable();

            try
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @strSql;
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

        //返回String
        public string GetStringbySql(string strSql, SqlConnection conn)
        {
            string retName = null;
            var sqlColumns = new List<object>();

            try
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @strSql;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                        while (reader.Read()) sqlColumns.Add(reader[0].ToString());
                }
            }
            catch (Exception ex)
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("GetStringbySql查询错误信息:" + ex.Message + "==SQL:" +
                                                                           strSql);
            }

            if (sqlColumns.Count != 0 && sqlColumns[0] != null)
            {
                retName = sqlColumns[0].ToString();
            }
            return retName;
        }

        public static bool BackupDatabase(DBConnection dbConnection, int nStepNum, SqlConnection conn)
        {
            var sBackupPath =
                $@"D:\BizSolution\DBbackup\{dbConnection.Database}_{DateTime.Now:yyyyMMdd}_Before{nStepNum}.bak";
            var sBackupFileName = Path.GetFileNameWithoutExtension(sBackupPath);
            return BackupDatabase(dbConnection, sBackupPath, conn);
        }

        public static bool BackupDatabase(DBConnection dbConnection, string sBackupFileDir, SqlConnection conn)
        {
            var sBackupSql = $"BACKUP DATABASE {dbConnection.Database} TO DISK = '{sBackupFileDir}'";
            try
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sBackupSql;
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Bizsolution.BasicFacility.Exceptions.ExceptionLog.LogTimer("错误信息:" + ex.Message);
            }
            return true;

            #region 此方法提示未注册组件

            //SQLDMO.Backup oBackup = new SQLDMO.BackupClass();
            //SQLDMO.SQLServer oSQLServer = new SQLDMO.SQLServerClass();
            //try
            //{

            //    oSQLServer.LoginSecure = false;
            //    //下面设置登录sql服务器的ip,登录名,登录密码
            //    oSQLServer.Connect();
            //    oBackup.Action = 0;
            //    //数据库名称:
            //    oBackup.Database = dbConnection.Database;
            //    //备份的路径
            //    oBackup.Files = sBackupFileDir;
            //    //备份的文件名
            //    oBackup.BackupSetName = sBackupName;
            //    oBackup.BackupSetDescription = "数据库备份";
            //    oBackup.Initialize = true;
            //    oBackup.SQLBackup(oSQLServer);
            //}
            //catch
            //{
            //    return false;
            //}
            //finally
            //{
            //    oSQLServer.DisConnect();
            //}
            //return true; 

            #endregion
        }
    }

    public class DBConnection
    {
        public string ServerName { get; set; }
        public string LoginName { get; set; }
        public string Password { get; set; }
        public string Database { get; set; }
        public string ConnectionString { get; set; }
    }
}
