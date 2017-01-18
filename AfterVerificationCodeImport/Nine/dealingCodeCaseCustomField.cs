using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport.Nine
{
    class dealingCodeCaseCustomField
    {
        private readonly DBHelper _dbHelper = new DBHelper();

        public int InsertCodeCaseCustomField(DataRow dataRow, int row, string commDB, SqlConnection _connection)
        {
            if (dataRow["自定义属性名称"] != null && !string.IsNullOrEmpty(dataRow["自定义属性名称"].ToString()))
            {
                string sNo = dataRow["我方卷号"].ToString().Trim();
                int numHk = _dbHelper.GetIDbyName(sNo, 2, _connection);
                if (numHk > 0)
                {
                    string strSql =
                        " SELECT n_ID FROM TCode_CaseCustomField WHERE  s_IPType='P' AND s_IsActive='Y' AND s_CustomFieldName IN ('" +
                        dataRow["自定义属性名称"] + "')";
                    int nCaseFieldID = _dbHelper.GetbySql(strSql, commDB, _connection);
                    if (nCaseFieldID > 0)
                    {
                        strSql = " SELECT n_ID FROM TCase_CaseCustomField WHERE n_CaseID=" + numHk +
                                 " AND n_FieldCodeID=" + nCaseFieldID;
                        int nID = _dbHelper.GetbySql(strSql, commDB, _connection);
                        if (nID > 0)
                        {
                            strSql = "update TCase_CaseCustomField set s_Value='" + dataRow["自定义属性值"] +
                                     "' WHERE n_CaseID=" +
                                     numHk + " AND n_FieldCodeID=" + nCaseFieldID;
                        }
                        else
                        {
                            strSql =
                                "INSERT INTO dbo.TCase_CaseCustomField( n_CaseID, n_FieldCodeID, s_Value ) VALUES  (" +
                                numHk + "," + nCaseFieldID + "," + dataRow["自定义属性值"] + ")";
                        }
                    }
                    else
                    {
                        _dbHelper.InsertLog(0, sNo, row, "自定义属性", "自定义属性-" + row, "未找到自定义属性-：" + sNo, "", commDB, _connection); 
                    }
                    return _dbHelper.InsertbySql(strSql, row, commDB, _connection);
                }
                else
                {
                    _dbHelper.InsertLog(0, sNo, row, "自定义属性", "自定义属性-" + row, "未找到“我方卷号”为：" + sNo, "", commDB, _connection); 
                }
            }
            return 0;
        }

        //专利数据补充
        public int InsertPantentCodeCaseCustomField(string codeName, string codeNameValue, string sNo, int row, string commDB, SqlConnection _connection)
        { 
                int numHk = _dbHelper.GetIDbyName(sNo, 2, _connection);
                if (numHk > 0)
                {
                    string strSql =
                        " SELECT n_ID FROM TCode_CaseCustomField WHERE  s_IPType='P' AND s_IsActive='Y' AND s_CustomFieldName IN ('" +
                        codeName + "')";
                    int nCaseFieldID = _dbHelper.GetbySql(strSql, commDB, _connection);
                    if (nCaseFieldID > 0)
                    {
                        strSql = " SELECT n_ID FROM TCase_CaseCustomField WHERE n_CaseID=" + numHk +
                                 " AND n_FieldCodeID=" + nCaseFieldID;
                        int nID = _dbHelper.GetbySql(strSql, commDB, _connection);
                        if (nID > 0)
                        {
                            strSql = "update TCase_CaseCustomField set s_Value='" + codeNameValue +
                                     "' WHERE n_CaseID=" +
                                     numHk + " AND n_FieldCodeID=" + nCaseFieldID;
                        }
                        else
                        {
                            strSql =
                                "INSERT INTO dbo.TCase_CaseCustomField( n_CaseID, n_FieldCodeID, s_Value ) VALUES  (" +
                                numHk + "," + nCaseFieldID + "," + codeNameValue + ")";
                        }
                    }
                    else
                    {
                        _dbHelper.InsertLog(0, sNo, row, "自定义属性-专利数据", "自定义属性-专利数据-" + row, "未找到“我方卷号”为：" + sNo, "", commDB, _connection); 
                    }
                    return _dbHelper.InsertbySql(strSql, row, commDB, _connection);
                
            }
            return 0;
        }

    }
}
