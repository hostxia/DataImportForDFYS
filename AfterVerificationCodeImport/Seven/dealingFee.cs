using System;
using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport.Seven
{
    class dealingFee
    {
        readonly DBHelper _dbHelper = new DBHelper();

        public int InsertFee(int rowid, DataRow dr,string commDB, SqlConnection _connection)
        {
            int result = 0;
            int year = 0;
            string sNo = dr["案件文号"].ToString().Trim();
            int hkNum = _dbHelper.GetIDbyName(sNo, 2,_connection);
            if (hkNum.Equals(0))
            {
                //未找到“我方卷号” 
                _dbHelper.InsertLog(0, sNo, rowid, "国内-年费", "国内-年费-" + rowid, "未找到“我方卷号”为：" + sNo, "", commDB, _connection);
                return result;
            }
            else
            {
                result = 1;
                //专利案件申请日  
                string Sql = "SELECT dt_AppDate FROM TCase_Base WHERE n_CaseID=" + hkNum;
                string time = _dbHelper.GetStringbySql(Sql, _connection);
                if (time != null && time.ToString() != "")
                {
                    year = DateTime.Parse(time.ToString()).Year;
                }
                 
                //专利类型
                Sql = "SELECT n_PatentTypeID FROM TPCase_Patent WHERE n_CaseID=" + hkNum;
                int n_PatentTypeID = _dbHelper.GetbySql(Sql, commDB, _connection);

                //读取年费标准 
                Sql = "SELECT  n_YearNo,n_OfficialFee FROM TCode_AnnualFee WHERE n_PatentType=" + n_PatentTypeID +
                      " ORDER BY n_YearNo asc";
                DataTable tableYearNo = _dbHelper.GetDataTablebySql(Sql,_connection);

                //读取年费标准年数
                Sql = " SELECT  count(*) as YearSum  FROM TCode_AnnualFee WHERE n_PatentType=" + n_PatentTypeID;
                int YearSum = _dbHelper.GetbySql(Sql, commDB, _connection);
                if (dr["下次年费年度"].ToString() != "")
                {
                    DateTime Next = DateTime.Parse(dr["下次年费年度"].ToString());
                    int NextTime = Next.Year;

                    int Sumnum = year + YearSum - NextTime;
                    int Start = YearSum - Sumnum + 1;
                    if (tableYearNo != null)
                    {
                        for (int iS = Start; iS < YearSum + 1; iS++)
                        {
                            //查询是否存在当前年份的年费
                            Sql = "SELECT n_AnnualFeeID FROM T_AnnualFee WHERE n_CaseID=" + hkNum + " AND n_YearNo=" + iS;
                            int n_AnnualFeeID = _dbHelper.GetbySql(Sql, commDB, _connection);
                            if (n_AnnualFeeID > 0)
                            {
                                //如果存在当前年的年费不做处理
                                Sql = "update T_AnnualFee set dt_OfficialShldPayDate='" + Next + "',dt_AlarmDate='" +
                                      Next.AddMonths(-2) + "' WHERE n_AnnualFeeID=" + n_AnnualFeeID;
                                int numS = _dbHelper.InsertbySql(Sql, rowid, commDB, _connection);
                                if (numS == 0)
                                {
                                    _dbHelper.InsertLog(hkNum, sNo, rowid, "国内-年费", "国内-年费-" + rowid, "修改年费数据插入错误：" + sNo, Sql.Replace("'", "''"), commDB, _connection); 
                                }
                            }
                            else
                            {
                                //循环产生年费记录
                                Sql =
                                    "INSERT INTO dbo.T_AnnualFee( n_CaseID ,n_YearNo , s_Status , s_PayMode , s_StatusOrder , n_ChargeCurrency , n_ChargeOFee , n_OfficialCurrency , n_OfficialFee ,  s_IsOfficialDisc ,s_OfficialDiscStyle , dt_OfficialShldPayDate ,dt_AlarmDate,s_IsActive ,dt_CreateDate ,dt_EditDate)" +
                                    "VALUES  (" + hkNum + "," + (iS) + ",'XXNNN','AX','123' ,8 ,'" +
                                    tableYearNo.Rows[iS - 1]["n_OfficialFee"] + "' ,8 ,'" +
                                    tableYearNo.Rows[iS - 1]["n_OfficialFee"] + "','N' ,'2' , '" + Next + "','" +
                                    Next.AddMonths(-2) + "','Y','" + DateTime.Now + "','" + DateTime.Now + "')";

                                int numS = _dbHelper.InsertbySql(Sql, rowid, commDB, _connection);
                                result = numS;
                                if (numS == 0)
                                {
                                    _dbHelper.InsertLog(hkNum, sNo, rowid, "国内-年费", "国内-年费-" + rowid, "增加年费数据插入错误：" + sNo, Sql.Replace("'", "''"), commDB, _connection); 
                                }
                            }
                            Next = Next.AddYears(1);
                        }
                    }
                }
                else
                {
                    _dbHelper.InsertLog(hkNum, sNo, rowid, "国内-年费", "国内-年费-" + rowid, "下次年费年度为空,无法导入：" + sNo, "", commDB, _connection); 
                }
            }
            return result;
        }

    }
}
