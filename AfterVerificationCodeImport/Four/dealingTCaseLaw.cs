using System;
using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport.Four
{
    class dealingTCaseLaw
    {
        private readonly DBHelper _dbHelper = new DBHelper();
        readonly T_MainFiles _tMainFiles=new T_MainFiles();
        public int TCaseLaw(int rowid, DataRow dr, int InFileID, int OutFileID, string commDB, SqlConnection _connection)
        { 
            string sNo = dr["我方卷号"].ToString().Trim();
            int hkNum = _dbHelper.GetIDbyName(sNo, 2, _connection);
           
            if (hkNum.Equals(0))
            {
                _dbHelper.InsertLog(0, sNo, rowid, "国外-法律信息及日志表", "国外-法律信息及日志表-" + rowid, "未找到“我方卷号”为：" + sNo, "", commDB, _connection);
                return 0;
            }
            DateTime insertTime = DateTime.Now.AddMonths(-1);
            if (dr["项目"] != null && dr["项目"].ToString().Trim() != "办登日" && dr["项目"].ToString().Trim() != "IDS号" &&
                dr["项目"].ToString().Trim() != "IDS日" && dr["项目"].ToString().Trim() != "驳回日" &&
                dr["项目"].ToString().Trim() != "复审日")
            {
                #region
                string sIoType = string.Empty; //文件类型
                string sClientGov = string.Empty; //客户 官方
                string type = string.Empty;
                int nAgencyID = 0;
                #region 判断来文类型

                if (dr["项目"] != null && dr["项目"].ToString().Trim() == "官方发文")//官方发文   		官方来文
                {
                    type = "来文";
                    sIoType = "I";
                    sClientGov = "C";
                }

                if (dr["项目"] != null &&
                   (dr["项目"].ToString().Trim() == "申请人来文" || dr["项目"].ToString().Trim() == "委托人来文"))// 申请人来文、委托人来文		    客户来文
                {
                    type = "来文";
                    sIoType = "I";
                    sClientGov = "O";
                }
                if (dr["项目"] != null &&
                    (dr["项目"].ToString().Trim() == "我方递交官方文件" || dr["项目"].ToString().Trim() == "我方给委托人文" || dr["项目"].ToString().Trim() == "我方给申请人来文"))//我方给申请人来文、我方给委托人文	发客户文
                { //我方递交官方文件	发官方文
                    type = "发文";
                    sIoType = "O";
                    sClientGov = dr["项目"].ToString().Trim() != "我方递交官方文件" ? "C" : "O";
                }
                string strSql;
                if (dr["项目"] != null && dr["项目"].ToString().Trim() == "我方给外代理文")//我方给外代理文   		发代理机构文
                {
                    type = "发文";
                    sIoType = "O";
                    sClientGov = "C";
                    strSql = " SELECT n_CoopAgencyToID  FROM dbo.TCase_Base WHERE n_CaseID=" + hkNum;
                    nAgencyID = _dbHelper.GetbySql(strSql, commDB, _connection);
                    //if (nAgencyID <= 0)
                    //{
                    //    _dbHelper.InsertLog(hkNum, sNo, rowid, "国外-法律信息及日志表", "国外-法律信息及日志表-" + rowid, "项目类型：我方给外代理文，未查到案件代理机构", "", commDB, _connection);
                    //}
                }
                string sClientType = "";
                string sNote = "";
                if (dr["项目"] != null && dr["项目"].ToString().Trim() == "外代理来文")//外代理来文   		外代理来文（系统目前不支持） 	
                {
                    type = "来文";
                    sIoType = "I";
                    sClientGov = "O";
                    sClientType = "O";
                    sNote = "[来文类型：外代理]";
                }
                #endregion

                #region
                strSql = "";
                if (type == "来文" || type == "发文")
                {

                    if (type == "来文")
                    {
                        strSql =
                            "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,s_Abstact,ObjectType,s_SendMethod,dt_EditDate,s_Name,dt_ReceiveDate,dt_CreateDate,s_IOType,s_ClientGov,s_Status,s_ClientType)" +
                            "VALUES  ('国外-法律信息及日志表" + rowid + "',''," + InFileID + ",'其他','" + insertTime + "','" +
                            dr["项目"].ToString().Replace("'", "''") + "  " +
                            dr["内容"].ToString().Replace("'", "''") + "','" + dr["记录日"] + "','" + insertTime + "','" +
                            sIoType + "','" + sClientGov + "','Y','" + sClientType + "')";
                    }
                    else if (type == "发文")
                    {
                        strSql =
                            "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,s_Abstact,ObjectType,s_SendMethod,dt_EditDate,s_Name,dt_SendDate,dt_CreateDate,s_IOType,s_ClientGov,s_Status)" +
                            "VALUES  ('国外-法律信息及日志表" + rowid + "',''," + OutFileID + ",'其他','" + insertTime + "','" +
                            dr["项目"].ToString().Replace("'", "''") + "  " +
                            dr["内容"].ToString().Replace("'", "''") + "','" + dr["记录日"] + "','" + insertTime + "','" +
                            sIoType + "','" + sClientGov + "','Y')   ";
                    }
                    if (!string.IsNullOrEmpty(strSql))
                    {
                        _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                    }
                    int insertnum = 0;
                    string strSqlS = "SELECT top 1  n_FileID FROM dbo.T_MainFiles WHERE s_Status='Y' AND  s_Name='" +
                                     dr["项目"].ToString().Replace("'", "''") + "  " +
                                     dr["内容"].ToString().Replace("'", "''") + "' and s_ClientGov='" + sClientGov +
                                     "' and s_IOType='" + sIoType + "' and dt_CreateDate='" + insertTime + "' ";
                    if (type == "来文")
                    {
                        strSqlS += " and dt_ReceiveDate='" + dr["记录日"] + "' and ObjectType=" + InFileID;
                    }
                    else if (type == "发文")
                    {
                        strSqlS += " and dt_SendDate='" + dr["记录日"] + "' and ObjectType=" + OutFileID;
                    }
                    int nFileID = _dbHelper.GetbySql(strSqlS + " order by n_FileID desc ", commDB, _connection);
                    if (nFileID > 0)
                    {
                        insertnum = nFileID;
                    }
                    if (insertnum > 0)
                    {
                        int nGovOfficeID = 0;
                        if (sClientGov.Equals("O")) //s_ClientGov C: 客户 O: 官方
                        {
                            nGovOfficeID = 21; //中国国家知识产权局
                        }
                        strSql = "SELECT COUNT(*) AS sumNum FROM dbo.T_FileInCase WHERE n_CaseID=" + hkNum +
                                 " AND n_FileID=" + insertnum;
                        int sumNumFileInCase = _dbHelper.GetbySql(strSql, commDB, _connection);
                        if (sumNumFileInCase <= 0)
                        {
                            strSql = "SELECT COUNT(*) AS sumNum FROM dbo.T_FileInCase WHERE n_FileID=" + insertnum +
                                     " AND n_CaseID=" + hkNum;
                            int sumNumC = _dbHelper.GetbySql(strSql, commDB, _connection);
                            if (sumNumC <= 0)
                            {
                                strSql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                                         "VALUES  (" + hkNum + " ," + insertnum + ",'Y')";
                                _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                            }
                        }

                        if (type == "来文")
                        {
                            strSql = "SELECT COUNT(*) AS sumNum FROM dbo.T_InFiles WHERE n_FileID=" + insertnum +
                                     " and n_GovOfficeID=" + nGovOfficeID +
                                     " and s_Distribute='Y'  and s_OFileStatus='N'";
                            int sumNum = _dbHelper.GetbySql(strSql, commDB, _connection);
                            if (sumNum <= 0)
                            {
                                strSql =
                                    "INSERT INTO dbo.T_InFiles( n_FileID,n_FileCodeID,n_GovOfficeID,s_OFileStatus,s_Distribute,s_Note)" +
                                    "VALUES  (" + insertnum + ",0," + nGovOfficeID + " ,'N','Y','" + sNote + "')";
                                _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                            }
                        }
                        else if (type == "发文")
                        {
                            strSql =
                                "INSERT INTO dbo.T_OutFiles( n_FileID ,n_CheckedOutBy , n_GovOfficeID , s_FileStatus, dt_StatusDate ,dt_WriteDate ," +
                                "n_WriterID , n_SubmiterID ,  n_PrintNum , n_PageNum ,n_ReFileID  ,n_Count ,s_FileType ,n_LatestCheckInfoID";
                            string drValue = "VALUES  (" + insertnum + ",0 ," + nGovOfficeID + " ,'W' ,'" + DateTime.Now + "' ,'" +
                              DateTime.Now + "',0 ,0 ,1,0 ,0,0 ,'new',0";
                            if (nAgencyID > 0)
                            {
                                strSql += ",n_AgencyID";
                                drValue += "," + nAgencyID;
                            }
                            strSql = strSql + ")   " + drValue + ")";
                            string sql = "SELECT COUNT(*) AS sumcount FROM dbo.T_OutFiles WHERE n_FileID=" +
                                         insertnum;
                            int sumcount = _dbHelper.GetbySql(sql, commDB, _connection);
                            if (sumcount <= 0)
                            {
                                _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                            }
                        }

                    }
                }
                //法律信息
                else if (dr["项目"] != null &&
                         (dr["项目"].ToString().Trim() == "PCT公开号" || dr["项目"].ToString().Trim() == "PCT公开日" ||
                          dr["项目"].ToString().Trim() == "PCT进入日" || dr["项目"].ToString().Trim() == "PCT申请号" ||
                          dr["项目"].ToString().Trim() == "PCT申请日" || dr["项目"].ToString().Trim() == "PCT办登日" ||
                          dr["项目"].ToString().Trim() == "公开号" || dr["项目"].ToString().Trim() == "公开日" ||
                          dr["项目"].ToString().Trim() == "进入实审日" || dr["项目"].ToString().Trim() == "授权公告号" ||
                          dr["项目"].ToString().Trim() == "授权公告日"))
                {
                    strSql = "UPDATE dbo.TPCase_LawInfo set " + GetColum(dr["项目"].ToString()) + "='" +
                             dr["内容"].ToString().Replace("'", "''") + "'";
                    strSql += "  WHERE n_ID IN (SELECT n_LawID FROM TPCase_Patent WHERE n_CaseID=" + hkNum + ")";
                    _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                }
                else if (dr["项目"] != null && dr["项目"].ToString().Trim() == "提实审日")
                {
                    strSql = "UPDATE dbo.TPCase_Patent set dt_RequestSubmitDate='" +
                             dr["内容"].ToString().Replace("'", "''") + "' WHERE n_CaseID=" + hkNum;
                    _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                }
                #endregion 
                return 1;
                #endregion
            }
            else
            {
                #region IDS号和IDS日

                if (dr["项目"] != null && (dr["项目"].ToString().Trim() == "IDS号" || dr["项目"].ToString().Trim() == "IDS日"))
                { 
                        string notes =_dbHelper.GetStringbySql("select s_Notes from TPCase_Patent WHERE n_CaseID=" + hkNum,_connection);
                        string content = "";
                        if (!string.IsNullOrEmpty(notes))
                        {
                            content = notes + "\r\n";
                        }
                        content += Environment.NewLine + "项目：" + dr["项目"] + " 内容:" + dr["内容"];
                        string strSql = " UPDATE TPCase_Patent SET s_Notes='" + content + "' WHERE n_CaseID=" + hkNum;

                        _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                        //IDS号、IDS日生成一条 发官方文 记录，IDS日为发文日，IDS号为发文名称。
                        string idsName = "";
                        string time = "1900-01-01 00:00:00.000";
                        if (dr["项目"] != null && dr["项目"].ToString().Trim() == "IDS号")
                        {
                            idsName = dr["内容"].ToString();
                        }
                        else if (dr["项目"] != null && dr["项目"].ToString().Trim() == "IDS日")
                        {
                            time = dr["内容"].ToString();
                        }

                        //查询文件主表是否包含此文件记录
                        string strSql1 = " SELECT n_FileID FROM dbo.T_MainFiles WHERE    s_Name='" + idsName +
                                         "' AND dt_SendDate='" + time + "' and dt_CreateDate='" + insertTime + "'";
                        int nFileID = _dbHelper.GetbySql(strSql1, commDB, _connection);
                        if (nFileID > 0)
                        {
                            string strSql2 = " SELECT n_FileID FROM dbo.T_OutFiles WHERE n_FileID=" + nFileID;
                            int nFileIDOut = _dbHelper.GetbySql(strSql2, commDB, _connection);
                            if (nFileIDOut > 0) //发件表和文件主表关联
                            {
                                string strSql3 = "  select n_ID FROM T_FileInCase  WHERE n_FileID =" + nFileID +
                                                 " AND n_CaseID=" + hkNum;
                                int nID = _dbHelper.GetbySql(strSql3, commDB, _connection);
                                if (nID <= 0) //发件和无案件关联
                                {
                                    _tMainFiles.InserIntoTFileInCase(hkNum, nFileID, rowid, sNo, "国外-法律信息及日志表", commDB, _connection);
                                } 
                            }
                            else
                            {
                                if (_tMainFiles.InserIntoTOutFiles(21, nFileID, rowid, commDB, _connection) > 0)
                                {
                                    _tMainFiles.InserIntoTFileInCase(hkNum, nFileID, rowid, sNo, "国外-法律信息及日志表", commDB, _connection);
                                }
                            }
                        }
                        else
                        {
                            if (_tMainFiles.InserIntoTMainFiles(idsName, insertTime, time, rowid, hkNum, sNo, OutFileID, "国外-法律信息及日志表", commDB, _connection) > 0)
                            {
                                nFileID = _dbHelper.GetbySql(strSql1, commDB, _connection);
                                if (nFileID > 0)
                                {
                                    if (_tMainFiles.InserIntoTOutFiles(21, nFileID, rowid, commDB, _connection) > 0)
                                    {
                                        string strSql2 = " SELECT n_FileID FROM dbo.T_OutFiles WHERE n_FileID=" + nFileID;
                                        int nFileIDOut = _dbHelper.GetbySql(strSql2, commDB, _connection);
                                        if (nFileIDOut > 0) //发件表和文件主表关联
                                        {
                                            _tMainFiles.InserIntoTFileInCase(hkNum, nFileID, rowid, sNo, "国外-法律信息及日志表", commDB, _connection);
                                        }
                                    }
                                }
                                else
                                {
                                    _dbHelper.InsertLog(hkNum, sNo, rowid, "国外-法律信息及日志表", "国外-法律信息及日志表-" + rowid, "未查询到文件数据：" + sNo, strSql1.Replace("'", "''"), commDB, _connection);
                                }
                            }
                            else
                            {
                                _dbHelper.InsertLog(hkNum, sNo, rowid, "国外-法律信息及日志表", "国外-法律信息及日志表-" + rowid, "增加数据失败T_MainFiles：" + sNo, "", commDB, _connection);
                            }
                        }
                    } 
                #endregion

                #region 办登日、复审日、驳回日
                else if (dr["项目"] != null &&
                        (dr["项目"].ToString().Trim() == "办登日" || dr["项目"].ToString().Trim() == "驳回日" ||
                         dr["项目"].ToString().Trim() == "复审日"))
                {
                    //导入到官方来文，生成来文记录，内容中的日期作为来文日期
                    string Name = dr["项目"].ToString().Replace("'", "''");
                    if (Name.Equals("办登日"))
                    {
                        Name = "办登通知书";
                    }
                    else if (Name.Equals("复审日"))
                    {
                        Name = "复审通知书";
                    }
                    else if (Name.Equals("驳回日"))
                    {
                        Name = "驳回决定通知书";
                    }
                    InserIntoTMainFilesIn(hkNum, Name, dr["内容"].ToString().Replace("'", "''"), insertTime, rowid, InFileID, commDB, _connection);
                }
                #endregion
                return 1;
            } 
        }

        #region 来文

        private void InserIntoTMainFilesIn(int hkNum, string project, string content, DateTime insertTime, int rowid, int InFileID, string commDB, SqlConnection _connection)
        {
            string strSql =
                "INSERT  INTO dbo.T_MainFiles(s_sourcetype1,s_Abstact,ObjectType,s_SendMethod,dt_EditDate,s_Name,dt_ReceiveDate,dt_CreateDate,s_IOType,s_ClientGov)" +
                "VALUES  ('国外-法律信息及日志表" + rowid + "',''," + InFileID + ",'其他','" + insertTime + "','" + project + "','" + content + "','" + insertTime +"','I','O')";
            _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
            strSql = "SELECT n_FileID FROM dbo.T_MainFiles WHERE   s_Name='" + project +
                     "' and s_ClientGov='O' and s_IOType='I'  and dt_ReceiveDate='" + content +
                     "' and ObjectType=" + InFileID + " and dt_CreateDate='" + insertTime + "'";
            int nFileID = _dbHelper.GetbySql(strSql, commDB, _connection);
            if (nFileID > 0)
            {
                strSql = "SELECT COUNT(*) AS sumNum FROM dbo.T_FileInCase WHERE n_FileID=" + nFileID + " AND n_CaseID=" +
                         hkNum;
                int sumNumC = _dbHelper.GetbySql(strSql, commDB, _connection);
                if (sumNumC <= 0)
                {
                    strSql = "INSERT INTO dbo.T_FileInCase(n_CaseID,n_FileID,s_IsMainCase)" +
                             "VALUES  (" + hkNum + " ," + nFileID + ",'Y')";
                    _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                }
                strSql = "SELECT COUNT(*) AS sumNum FROM dbo.T_InFiles WHERE n_FileID=" + nFileID +
                         " AND n_GovOfficeID=21";
                int sumNum = _dbHelper.GetbySql(strSql, commDB, _connection);
                if (sumNum <= 0)
                {
                    strSql =
                        "INSERT INTO dbo.T_InFiles( n_FileID,n_FileCodeID,n_GovOfficeID,s_OFileStatus,s_Distribute)" +
                        "VALUES  (" + nFileID + ",0,21,'N','Y')";
                    _dbHelper.InsertbySql(strSql, rowid, commDB, _connection);
                }
            }
            else
            {
                _dbHelper.InsertLog(hkNum, "", rowid, "国外-法律信息及日志表", "国外-法律信息及日志表-" + rowid, "增加数据失败InserIntoTMainFilesIn", strSql.Replace("'", "''"), commDB, _connection);
            }
        }

        #endregion

        private static string GetColum(string name)
        {
            string strSql = "";
            switch (name)
            {
                case "PCT公开号":
                    strSql = "s_PCTPubNo";
                    break;
                case "PCT公开日":
                    strSql = "dt_PctPubDate";
                    break;
                case "PCT进入日":
                    strSql = "dt_PctInNationDate";
                    break;
                case "PCT申请号":
                    strSql = "s_PCTAppNo";
                    break;
                case "PCT申请日":
                    strSql = "dt_PctAppDate";
                    break;
                case "PCT办登日":
                    strSql = "dt_CertfDate";
                    break;
                case "公开号":
                    strSql = "s_PubNo";
                    break;
                case "公开日":
                    strSql = "dt_PubDate";
                    break;
                case "进入实审日":
                    strSql = "dt_SubstantiveExamDate";
                    break;
                case "授权公告号":
                    strSql = "s_IssuedPubNo";
                    break;
                case "授权公告日":
                    strSql = "dt_IssuedPubDate";
                    break;
            }
            return strSql;
        }
    }
}
