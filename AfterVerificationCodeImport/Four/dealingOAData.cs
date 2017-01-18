using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport.Four
{
    internal class dealingOAData
    {
        private readonly DBHelper _dbHelper = new DBHelper();

        //OA数据补充导入
        public int OAData(int _rowid, DataRow dr, int OutFileID, int InFileID, string commDB, SqlConnection _connection)
        {
            int result = 0;
            string s_No = dr["我方卷号"].ToString().Trim();
            int hkNum = _dbHelper.GetIDbyName(s_No, 2, _connection);
            if (hkNum.Equals(0))
            {
                _dbHelper.InsertLog(0, s_No, _rowid, "国内-OA数据补充导入", "国内-OA数据补充导入-" + _rowid, "未找到“我方卷号”为：" + s_No, "",
                                    commDB, _connection);
                return result;
            }
            if (hkNum != 0)
            {
                string nameType;
                string sClientGov;
                if (!string.IsNullOrEmpty(dr["转发客户日"].ToString().Trim()))
                {
                    nameType = dr["类型"] == null ? "OA转发客户" : dr["类型"] + "转客户";
                    sClientGov = "C";
                    InsertIntos(_rowid, dr, sClientGov, nameType, hkNum, dr["转发客户日"].ToString().Trim(), OutFileID, commDB, _connection);
                }
                if (!string.IsNullOrEmpty(dr["账单及代理报告发出日"].ToString().Trim()))
                {
                    nameType = "报告及账单发客户";
                    sClientGov = "C";
                    InsertIntos(_rowid, dr, sClientGov, nameType, hkNum, dr["账单及代理报告发出日"].ToString().Trim(), OutFileID, commDB, _connection);
                }
                if (!string.IsNullOrEmpty(dr["答复日"].ToString().Trim()))
                {
                    nameType = "发官方文";
                    sClientGov = "O";
                    InsertIntos(_rowid, dr, sClientGov, nameType, hkNum, dr["答复日"].ToString().Trim(), OutFileID, commDB, _connection);
                }
                if (!string.IsNullOrEmpty(dr["OA收到日"].ToString().Trim()))
                {
                    nameType = "官方来文";
                    sClientGov = "O";
                    InsertIntos(_rowid, dr, sClientGov, nameType, hkNum, dr["OA收到日"].ToString().Trim(), InFileID, commDB, _connection);
                }
                result = 1;
            }
            return result;
        }

        //来文和发文记录
        private void InsertIntos(int rowID, DataRow dr, string sClientGov, string nameType, int numHk, string time,
                                 int OutFileID, string commDB, SqlConnection _connection)
        {
            var _tMainFiles = new T_MainFiles();
            string remark = dr["OA延期"] == null ? "" : dr["OA延期"].ToString();
            string people = dr["代理人"] == null ? "" : dr["代理人"].ToString();

            remark = remark == "" ? "" : "OA延期:" + remark;
            people = people == "" ? "" : "代理人:" + people;
            remark = remark.Replace("'", "''") + "  " + people.Replace("'", "''");

            _tMainFiles.OADataInsertTMainFile(numHk, rowID, nameType, "国内-OA数据补充导入", time, OutFileID, remark, sClientGov,
                                              "O", commDB, _connection);
        }
   
    
        //案件处理人
        public int InsertCaseAttorney(DataTable table, int _rowid, DataRow dr, string commDB, SqlConnection _connection)
        {
            int result = 0;
            string sNo = dr["我方卷号"].ToString().Trim();
            int hkNum = _dbHelper.GetIDbyName(sNo, 2, _connection);
            if (hkNum.Equals(0))
            {
                _dbHelper.InsertLog(0, sNo, _rowid, "国内-OA数据补充导入(辅表)-已翻译代理人", "国内-OA数据补充导入(辅表)-已翻译代理人-" + _rowid, "未找到“我方卷号”为：" + sNo, "",
                                                  commDB, _connection);
                return 0;
            }
            else
            {
                result = 1;
                if (!string.IsNullOrEmpty(dr["代理人"].ToString().Trim()))
                {
                    string[] ArryUser = dr["代理人"].ToString().Split(' ');
                    if (ArryUser.Length > 0)
                    {
                        if (!string.IsNullOrEmpty(ArryUser[0]))
                        {
                            var _dealingCasePantent = new dealingCasePantent();
                            _dealingCasePantent.InsertUser(ArryUser[0].Trim().Replace("　", ""), hkNum, _rowid, "代理部-OA阶段-办案人", "国内-OA数据补充导入(辅表)-已翻译代理人", commDB, _connection);
                        }
                    }
                }
            }
            return result;
        }
    }
}
