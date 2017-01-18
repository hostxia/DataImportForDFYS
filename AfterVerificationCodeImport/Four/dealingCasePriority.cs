using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport.Four
{
    class dealingCasePriority
    {
        readonly DBHelper _dbHelper = new DBHelper();

        //国内优先权
        public int TPCasePriority(int _rowid, DataRow dr,string commDB, SqlConnection _connection)
        {
            int result = 0;
            #region
            string sNo = dr["我方卷号"].ToString().Trim();
            int Country = _dbHelper.GetIDbyName(dr["优先权国家"].ToString().Trim(), 1,_connection);
            int nCaseID = _dbHelper.GetIDbyName(sNo, 2, _connection);
            if (nCaseID.Equals(0))
            {
                _dbHelper.InsertLog(0, sNo, _rowid, "国内-优先权", "国内-优先权-" + _rowid, "未找到“我方卷号”为：" + sNo, "",
                                                  commDB,_connection); return result;
            }
            else
            {
                result = 1;
                InsertTPCasePriority(nCaseID, Country, dr, "国内-优先权", _rowid, commDB, _connection);
            }
            UpdateSeq(nCaseID, commDB, _connection);

            if (dr["优先权国家"].ToString().Trim().Equals("中国")) //B为主案
            {
                var _tCaseRelative = new TCaseRelative();
                _tCaseRelative.InsertInto(dr["优先权号"].ToString().Trim(), nCaseID, _rowid, "国内-优先权", commDB, _connection);
            }
            #endregion
            return result;
        }

        //增加优先权
        private void InsertTPCasePriority(int HKNum, int Country, DataRow dr, string excelName, int rowid, string commDB, SqlConnection _connection)
        {
            string strSql = "select n_ID from TPCase_Priority WHERE n_CaseID=" + HKNum + " AND s_PNum='" +
                                dr["优先权号"].ToString().Trim() + "'" +
                                " and n_PCountry=" + Country +
                                " and s_PDocProvided='N' and s_PTransDocProvided='N'";
            if (!string.IsNullOrEmpty(dr["优先权日"].ToString().Trim()))
            {
                strSql += " and dt_PDate='" + dr["优先权日"].ToString().Trim() + "'";
            }
            if (_dbHelper.GetbySql(strSql, commDB, _connection) <= 0)
            {
                strSql =
                    "INSERT INTO dbo.TPCase_Priority( n_CaseID ,n_Sequence , n_PCountry ,s_PNum ,s_PDocProvided , s_PTransDocProvided";
                string strsql2 = "VALUES  ('" + HKNum + "',100," + Country + ",'" +
                                 dr["优先权号"].ToString().Trim() + "','N','N'";
                if (!string.IsNullOrEmpty(dr["优先权日"].ToString().Trim()))
                {
                    strSql += " , dt_PDate";
                    strsql2 += ",'" + dr["优先权日"].ToString().Trim() + "'";
                }
                strSql += ")";
                strsql2 += ")";
                string sql = strSql + strsql2;
                if (_dbHelper.InsertbySql(sql, rowid, commDB, _connection) <= 0)
                {
                    _dbHelper.InsertLog(HKNum, "", rowid, "国内-优先权", "国内-优先权-" + rowid, "增加优先权失败", sql.Replace("'", "''"), commDB, _connection); 
                }
            }
            else
            {
                _dbHelper.InsertLog(HKNum, "", rowid, "国内-优先权", "国内-优先权-" + rowid, "案件已存在此优先权", strSql.Replace("'", "''"), commDB, _connection); 
            }
        }
      
        //重新排列优先权顺序
        private void UpdateSeq(int nCaseID, string commDB, SqlConnection _connection)
        {
            string strSql = " SELECT  * FROM  TPCase_Priority WHERE n_CaseID=" + nCaseID + " ORDER BY dt_PDate ASC  ";
            DataTable table = _dbHelper.GetDataTablebySql(strSql, _connection);
            for (int i = 0; i < table.Rows.Count; i++)
            {
                strSql = "update TPCase_Priority set n_Sequence=" + i + " where n_ID=" + table.Rows[i]["n_ID"];
                _dbHelper.InsertbySql(strSql, 0, commDB, _connection);
            }
        }

        //香港优先权
        public int HongKangPriority(int rowid, DataRow dr, string commDB, SqlConnection _connection)
        { 
            int Country = _dbHelper.GetIDbyName(dr["优先权国家"].ToString().Trim(), 1,_connection);
            int HKNum = _dbHelper.GetIDbyName(dr["我方卷号"].ToString().Trim(), 2, _connection);
            if (HKNum.Equals(0))
            {
                _dbHelper.InsertLog(0, dr["我方卷号"].ToString().Trim(), rowid, "香港-优先权", "香港-优先权-" + rowid, "未找到“我方卷号”为：" + dr["我方卷号"].ToString(), "", commDB, _connection); 
                return 0;
            }
            else
            {
                string strSql = "SELECT a.n_CaseID FROM tcase_Base  a  left join  TPCase_Patent b on a.n_CaseID=b.n_CaseID left join  TPCase_OrigPatInfo  c on b.n_OrigPatInfoID=c.n_ID " +
                              "where c.s_CaseSerial='" + dr["我方卷号"].ToString().Trim() + "'";

                DataTable table = _dbHelper.GetDataTablebySql(strSql,_connection);
                if (table.Rows.Count > 0)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        int HkNu = int.Parse(table.Rows[i]["n_CaseID"].ToString());
                        InsertTPCasePriority(HkNu, Country, dr, "香港-优先权", rowid, commDB, _connection);
                        if (dr["优先权国家"].ToString().Trim().Equals("中国")) //B为主案
                        {
                            var _tCaseRelative = new TCaseRelative();
                            _tCaseRelative.InsertInto(dr["优先权号"].ToString().Trim(), HKNum, rowid, "香港-优先权", commDB,
                                                      _connection);
                        }
                        UpdateSeq(HKNum, commDB, _connection);
                    }
                    return 1;
                }
                else
                {
                    _dbHelper.InsertLog(HKNum, "", rowid, "香港-优先权", "香港-优先权-" + rowid, "未查到香港案件的原案信息，无法添加优先权信息", strSql.Replace("'", "''"), commDB, _connection);
                }
                return 0;
            }
        }

    }
}
