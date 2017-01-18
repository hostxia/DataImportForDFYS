using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport.Four
{
    class dealingCaseInventor
    {
        private readonly DBHelper _dbHelper = new DBHelper();

        public int TPCaseInventor(int rowid, DataRow dr, string commDB, SqlConnection _connection)
        {
            string sCountry = dr["发明人国籍"].ToString().Trim();
            dr["发明人国籍"] = _dbHelper.GetIDbyName(sCountry, 1,_connection);

            string sNo = dr["我方卷号"].ToString().Trim();
            int hkNum = _dbHelper.GetIDbyName(sNo, 2, _connection);
            if (hkNum.Equals(0))
            {
                //未找到“我方卷号” 
                _dbHelper.InsertLog(0, sNo, rowid, "国外-国外库发明人表", "国外-国外库发明人表-" + rowid, "未找到“我方卷号”为：" + sNo, "", commDB, _connection);
                return 0;
            }
            if (hkNum != 0)
            {
                if (!string.IsNullOrEmpty(dr["发明人中文名"].ToString().Replace("'", "''")) || !string.IsNullOrEmpty(dr["发明人英文名"].ToString().Replace("'", "''")))
                {
                    string strSql = "SELECT n_ID FROM TPCase_Inventor WHERE n_CaseID=" + hkNum + " AND s_Name ='" +
                                    dr["发明人中文名"].ToString().Replace("'", "''") + "' AND s_NativeName='" +
                                    dr["发明人英文名"].ToString().Replace("'", "''") + "'";
                    int ResultNum = _dbHelper.GetbySql(strSql, commDB, _connection);
                    if (ResultNum <= 0)
                    {
                        int MaxSeq =
                            _dbHelper.GetbySql(
                                "SELECT TOP 1 n_Sequence  FROM TPCase_Inventor WHERE n_CaseID=" + hkNum +
                                " ORDER BY n_Sequence DESC ", commDB, _connection);
                        strSql =
                            " INSERT INTO dbo.TPCase_Inventor(n_Sequence,n_CaseID,s_NativeName ,s_Name,n_Country,s_Address)" +
                            " VALUES(" + MaxSeq + "," + hkNum + ",'" + dr["发明人英文名"].ToString().Replace("'", "''") +
                            "','" +
                            dr["发明人中文名"].ToString().Replace("'", "''") + "','" + dr["发明人国籍"] + "','" +
                            dr["发明人地址"].ToString().Replace("'", "''") + "')";
                    }
                    else
                    {
                        strSql = "update TPCase_Inventor set  s_Address='" + dr["发明人地址"].ToString().Replace("'", "''") +
                                 "'";
                        if (dr["发明人国籍"] != null && !string.IsNullOrEmpty(dr["发明人国籍"].ToString()) &&
                            dr["发明人国籍"].ToString().Trim() != "0")
                        {
                            strSql += ",n_Country=" + dr["发明人国籍"].ToString();
                        }
                        strSql += " where n_ID=" + ResultNum;
                    }
                    int iNum = _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                    if (iNum <= 0)
                    {
                        _dbHelper.InsertLog(0, sNo, rowid, "国外-国外库发明人表", "国外-国外库发明人表-" + rowid, "TPCaseInventor导入数据错误",
                                            strSql.Replace("'", "''"), commDB, _connection);
                    }
                    return iNum;
                }
                else
                {
                    _dbHelper.InsertLog(hkNum, sNo, rowid, "国外-国外库发明人表", "国外-国外库发明人表-" + rowid, "发明人信息为空", "", commDB, _connection); 
                    return 0;
                }
            }
            return 0;
        }

    }
}
