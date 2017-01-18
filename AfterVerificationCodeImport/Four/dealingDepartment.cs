using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AfterVerificationCodeImport.Four
{
    class dealingDepartment
    {
        readonly DBHelper _dbHelper = new DBHelper();
        public int UpdateOrg(DataRow row, int rowid,string commDB, SqlConnection _connection)
        {
            int result = 0;
            string sNo = row["客户"].ToString();
            string department = row["部门"].ToString();

            //查找部门ID
            string strSql = "select n_ID from T_Department where s_Name='" + department + "'";
            int Num = _dbHelper.GetbySql(strSql, commDB, _connection);
            if (Num <= 0)
            {
                string insql = "insert into T_Department(s_Name) values('" + department + "')";
                _dbHelper.InsertbySql(insql, rowid, commDB,_connection);
                Num = _dbHelper.GetbySql(strSql, commDB, _connection);
            }
            //查找案件客户
            if (Num > 0)
            {
                int nuA = _dbHelper.GetbySql("select count(n_CaseID) as sumNum from tcase_base  where  right(s_caseserial,3)='" + sNo, commDB,_connection);
                if (nuA > 0)
                {
                    strSql = " UPDATE TCase_Base SET n_DepartmentID=" + Num + " WHERE n_CaseID IN (select n_CaseID from tcase_base  where  right(s_caseserial,3)='" + sNo + "')";
                    result = _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                    if (result <= 0)
                    {
                        _dbHelper.InsertLog(0, "", rowid, "代理人部门", "代理人部门-" + rowid, "更新案件部门失败", strSql.Replace("'", "''"), commDB, _connection);
                    }
                }
                else
                {
                    _dbHelper.InsertLog(0, "", rowid, "代理人部门", "代理人部门-" + rowid, "未查到代理人部门案件", strSql.Replace("'", "''"), commDB, _connection);
                    result = 1;
                }
            }
            else
            {
                _dbHelper.InsertLog(0, "", rowid, "代理人部门", "代理人部门-" + rowid, "代理人部门未查到" + department, strSql.Replace("'","''"),commDB, _connection);
            }
            return result;
        }
    }
}
