using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport.Four
{
    class dealingTCodeBusinessType
    {
        readonly DBHelper _dbHelper = new DBHelper();

        public int UpdateType(DataRow row, int rowid, string commDB, SqlConnection _connection)
        {
            const int result = 0;
            string type = row["名称"].ToString();
            if (!string.IsNullOrEmpty(type))
            {
                string strSql = "SELECT n_ID FROM TCode_BusinessType  WHERE s_Name='" + type + "' and  s_IPType='P'";
                int nID = _dbHelper.GetbySql(strSql, commDB, _connection);

                string type1 = row["申请方式"].ToString().Trim();
                if (type1.Equals("纸件"))
                {
                    type1 = "N";
                }
                else if (type1.Equals("电子"))
                {
                    type1 = "Y";
                }
                else if (type1.Equals("不提交"))
                {
                    type1 = "U";
                }
                if (nID > 0)
                {
                    strSql = " UPDATE TCase_Base SET s_IsRegOnline='" + type1 + "' WHERE n_BusinessTypeID=" + nID;
                    return _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                }
                else
                {
                    _dbHelper.InsertLog(0, "", rowid, "申请方式", "申请方式-" + rowid, "申请方式为:" + type + "  业务类型ID：" + nID, "", commDB, _connection);
                }
            }
            return result;
        }

        public int UpdateUSAType(DataRow row, int rowid, string commDB, SqlConnection _connection)
        {
            string sNo = row["我方文号"].ToString();
            int HkNum = _dbHelper.GetIDbyName(sNo, 2, _connection);
            if (HkNum > 0)
            {
                string type = row["申请方式"].ToString().Trim();
                if (type.ToUpper().Equals("CA申请"))
                {
                    type = "A";
                }
                else if (type.ToUpper().Equals("CIP申请"))
                {
                    type = "P";
                }
                string Sql = "UPDATE TCase_Base SET s_IsRegOnline='" + type + "' WHERE n_CaseID=" + HkNum;
                return _dbHelper.InsertbySql(Sql, rowid, commDB, _connection);
            }
            else
            {
                _dbHelper.InsertLog(HkNum, sNo, rowid, "申请方式-美国", "申请方式-美国-" + rowid, "申请方式为:" + row["申请方式"] + "  不存在此案件：" + sNo, "", commDB, _connection);
            }
            return 0;
        }
    }
}
