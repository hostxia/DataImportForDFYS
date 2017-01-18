using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport.Four
{
    class dealingCaseToCase
    {
        readonly DBHelper _dbHelper = new DBHelper();

        public int InsertDoubleShen(DataRow dr, int rowid, string commDB, SqlConnection _connection)
        {
            int result = 0;
            string sNo = dr["我方卷号1（发明）"].ToString().Trim();
            int HKNum = _dbHelper.GetIDbyName(sNo, 2,_connection);
            int DoubleShenID =_dbHelper. GetIDbyName(dr["我方卷号2（实用新型）"].ToString().Trim(), 2,_connection);
            if (HKNum.Equals(0))
            {
                _dbHelper.InsertLog(0, sNo, rowid, "相关案件-双申", "相关案件-双申-" + rowid, "不存在我方卷号1（发明）：" + sNo, "", commDB, _connection);
                return result;
            }
            else
            {
                if (DoubleShenID.Equals(0))
                {
                    _dbHelper.InsertLog(0, sNo, rowid, "相关案件-双申", "相关案件-双申-" + rowid, "不存在我方卷号2（实用新型）：" + sNo, "", commDB, _connection); 
                }
                else
                {
                    const string strSql = "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='双申' AND s_MasterName='发明' AND s_SlaveName='实用新型' AND s_IPType='P'";
                    int n_ID = _dbHelper.GetbySql(strSql, commDB, _connection);
                    InsertIntoLaw(HKNum, DoubleShenID, n_ID, rowid, "", commDB, _connection);
                    result = 1;
                }
            }
            return result;
        }
        public void InsertIntoLaw(int nCaseID, int HKNum, int n_ID, int rowid, string type, string commDB, SqlConnection _connection)
        {
            int NUMS =
                _dbHelper.GetbySql("SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where n_CaseIDA=" + HKNum +
                           " and n_CaseIDB=" + nCaseID + " and n_CodeRelativeID=" + n_ID, commDB, _connection);
            if (!string.IsNullOrEmpty(type)) //同族
            {
                NUMS =
                     _dbHelper.GetbySql("SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where ((n_CaseIDA=" + HKNum +
                               " and n_CaseIDB=" + nCaseID + ") or (n_CaseIDA=" + nCaseID + " and n_CaseIDB=" + HKNum +
                               ") )and n_CodeRelativeID=" + n_ID, commDB, _connection);
            }

            if (NUMS <= 0 && HKNum > 0)
            {
                var _tMainFiles = new T_MainFiles();
                _tMainFiles.InsertTCaseCaseRelative(HKNum, nCaseID, n_ID, 0, rowid, commDB, _connection);
            }
        }

    }
}
