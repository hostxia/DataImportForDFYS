using System;
using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport.Seven
{
    class dealingTask
    {
        readonly DBHelper _dbHelper = new DBHelper();

        public int ImportTask(int rowid, DataRow dr,string commDB, SqlConnection _connection)
        {
            int resultNum = 0;
            string sNo = dr["案件文号"].ToString().Trim();
            int hkNum = _dbHelper.GetIDbyName(sNo, 2,_connection);
            if (hkNum.Equals(0))
            {
                _dbHelper.InsertLog(0, sNo, rowid, "任务时限", "任务时限-" + rowid, "未找到“我方卷号”为：" + sNo, "", commDB, _connection);
                return resultNum;
            }
            if (hkNum != 0)
            {
                string name = dr["时限名称"].ToString();
                string time = dr["时限日期"].ToString();
                
                string strSql = "SELECT n_ID FROM dbo.TCode_Employee WHERE s_Name='" + dr["负责人"].ToString() + "'";
                int nID = _dbHelper.GetbySql(strSql, commDB, _connection);
              
                //增加任务链
                string codeDeadlinegid = InsertTFCodeDeadline(name, rowid, commDB, _connection);
                //增加任务
                string taskgid = InsertTFTask(name, rowid, time, nID, commDB, _connection);
                //增加TFTaskChain
                string taskChaingid = InsertTFTaskChain(name, rowid, hkNum, sNo, commDB, _connection);
                //增加任务链节点
                InsertTFNode(taskChaingid, taskgid, rowid, commDB, _connection);
                //增加时限日期
                InsertTFDeadline(codeDeadlinegid, rowid, time, hkNum, taskChaingid, commDB, _connection);
                resultNum = 1;
            }
            return resultNum;
        }

        //任务时限名称
        private string InsertTFCodeDeadline(string Name, int rowid, string commDB, SqlConnection _connection)
        {
            string gid  = _dbHelper.GetStringbySql("select g_ID from TFCode_Deadline where s_Name='" + Name + "' and s_Type=''",_connection);
            if (string.IsNullOrEmpty(gid))
            {
                Guid guid = Guid.NewGuid();
                string strSql = "INSERT INTO  dbo.TFCode_Deadline( g_ID , s_Name,s_Type)" +
                                "VALUES  ('" + guid + "' ,'" + Name + "','' )";
                if (_dbHelper.InsertbySql(strSql, rowid, commDB, _connection) > 0)
                {
                    gid = guid.ToString();
                }
            } 
            return gid;
        }

        //任务
        private string InsertTFTask(string name, int rowid, string time, int peopel, string commDB, SqlConnection _connection)
        {
            string gid = "";
            Guid guid = Guid.NewGuid();
            string strSql =
                " INSERT INTO dbo.TF_Task(g_ID ,s_Name ,s_State ,s_ReadState ,dt_CreateTime ,dt_EditTime,n_ExecutorID";
            //)" +
            string strSql2 = "VALUES  ('" + guid + "' ,'旧任务-" + name + "','P' ,'R','" + DateTime.Now + "','" +
                             DateTime.Now + "'," + peopel;
            if (!string.IsNullOrEmpty(time))
            {
                strSql += ",dt_StartDate ,dt_EndDate";
                strSql2 += ",'" + time + "' ,'" + time + "'";
            }
            strSql += ")";
            strSql2 += ")";
            if (_dbHelper.InsertbySql(strSql + strSql2, rowid, commDB, _connection) > 0)
            {
                gid = guid.ToString();
            }
            return gid;
        }

        private string InsertTFTaskChain(string Name, int rowid, int CaseID, string s_CaseSerial, string commDB, SqlConnection _connection)
        {
            string gid = "";
            string sRelatedInfo2 =
                _dbHelper.GetStringbySql(
                    " SELECT TOP 1 '申请号：'+s_AppNo+';案件名称:'+s_CaseName AS s_RelatedInfo1 FROM dbo.TCase_Base  WHERE n_CaseID=" +
                    CaseID,_connection);
          
            Guid guid = Guid.NewGuid();
            string strSql =
                "  INSERT INTO dbo.TF_TaskChain( g_ID ,s_Name ,s_State ,s_TriggerType ,s_RelatedObjectType ,n_RelatedObjectID ,s_RelatedInfo1 , s_RelatedInfo2 ,dt_CreateTime ,dt_EditTime)" +
                "VALUES  ('" + guid + "' ,'旧任务链-" + Name + "','P' ,'Manual' ,'Case'," + CaseID + ",'" + s_CaseSerial +
                "','" + sRelatedInfo2 + "','" + DateTime.Now + "','" + DateTime.Now + "')";
            if (_dbHelper.InsertbySql(strSql, rowid, commDB, _connection) > 0)
            {
                gid = guid.ToString();
            }
            return gid;
        }

        //任务链节点  
        private void InsertTFNode(string g_TaskChainGuid, string g_FormerNodeGuid, int rowid, string commDB, SqlConnection _connection)
        {
            if (!string.IsNullOrEmpty(g_FormerNodeGuid))
            {
                Guid guid = Guid.NewGuid();
                string strSql = " INSERT INTO dbo.TF_Node( g_ID ,g_TaskChainGuid ,s_Mode ,s_Type )" +
                                "VALUES  ('" + guid + "' ,'" + g_TaskChainGuid + "','N' ,'S')";
                if (_dbHelper.InsertbySql(strSql, rowid, commDB, _connection) > 0) //开始
                {
                    Guid guid2 = Guid.NewGuid();
                    strSql =
                        " INSERT INTO dbo.TF_Node( g_ID ,g_TaskChainGuid ,g_FormerNodeGuid ,s_Mode ,s_Type , g_OwnTaskGuid)" +
                        "VALUES  ('" + guid2 + "' ,'" + g_TaskChainGuid + "','" + guid + "','N' ,'T' ,'" +
                        g_FormerNodeGuid + "')";
                    _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                }
                else
                {
                    _dbHelper.InsertLog(0, "", rowid, "任务时限", "任务时限-" + rowid, "查询错误信息[InsertTFNode]", strSql.Replace("'", "''"), commDB, _connection); 
                }
            }
        }

        //时限日期
        private void InsertTFDeadline(string g_CodeDeadlineID, int rowid, string Time, int CaseID, string TaskChaingid, string commDB, SqlConnection _connection)
        {
            string strSql =
                " INSERT INTO dbo.TF_Deadline(g_CodeDeadlineID ,s_RelatedObjectType ,n_RelatedObjectID ,dt_Deadline)" +
                "VALUES  ('" + g_CodeDeadlineID + "','Case'," + CaseID + ",'" + Time + "')";
            _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
             
            string gCodeTaskChainGuid =
                _dbHelper.GetStringbySql(" SELECT g_ID FROM dbo.TFCode_TaskChain WHERE   s_Code='LS001'", _connection);
            Guid guid = Guid.NewGuid();
            strSql = " INSERT INTO  dbo.TFCode_DeadlineInCodeTaskChain( g_ID ,g_CodeTaskChainGuid ,g_CodeDeadlineID)" +
                     "VALUES  ( '" + guid + "','" + gCodeTaskChainGuid + "' , '" + g_CodeDeadlineID + "' )";
            if (_dbHelper.InsertbySql(strSql, rowid, commDB, _connection) > 0)
            {
                strSql = "update TF_TaskChain set g_CodeTaskChainGuid='" + gCodeTaskChainGuid + "' where g_ID='" +
                         TaskChaingid + "'";
                _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
            }
        } 
    }
}
