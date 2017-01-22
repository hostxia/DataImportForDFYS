using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AfterVerificationCodeImport
{
    class TCaseRelative
    {
        private readonly DBHelper _dbHelper = new DBHelper();

        //添加国内优先权相关案件
        public void InsertInto(string No, int HKNum, int rowid, string TabName, string commDB, SqlConnection _connection)
        {
            string strSql = "SELECT n_ID FROM dbo.TCode_CaseRelative WHERE s_RelateName='国内优先权' AND s_MasterName='国内案' AND s_SlaveName='国外案' AND s_IPType='P'";
            int n_ID = _dbHelper.GetbySql(strSql, commDB, _connection);
               
            int caseID = _dbHelper.GetIDbyName(No, 7,_connection);//根据申请号查找案件
            if (caseID > 0)
            {
                strSql = "SELECT COUNT(*) AS SUM FROM dbo.TCase_CaseRelative where n_CaseIDA=" + HKNum + " and n_CaseIDB=" + caseID + " and n_CodeRelativeID=" + n_ID;
                int NUMS = _dbHelper.GetbySql(strSql, commDB, _connection);
                if (NUMS <= 0 && HKNum > 0)
                {
                    strSql = " INSERT INTO  dbo.TCase_CaseRelative ( n_CaseIDA ,  n_CaseIDB , dt_CreateDate , dt_EditDate , s_MasterSlaveRelation , n_CodeRelativeID )" +
                          " VALUES  ( " + HKNum + " , " + caseID + " ,  GETDATE() , GETDATE() ,  0,  " + n_ID + ")";
                    if (_dbHelper.InsertbySql(strSql, rowid, commDB, _connection) <= 0)
                    {
                        _dbHelper.InsertLog(HKNum, caseID.ToString(), rowid, TabName, TabName + rowid, "案件关系插入数据失败", strSql.Replace("'", "''"), commDB, _connection);
                    }
                }
            }
            //else
            //{
            //    _dbHelper.InsertLog(HKNum, caseID.ToString(), rowid, TabName, TabName + rowid, "未查到优先权号：无法建立优先权关系",No, commDB, _connection);
            //}
        }
    }
}
