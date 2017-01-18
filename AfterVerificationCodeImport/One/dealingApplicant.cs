using System;
using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport
{
    class dealingApplicant
    {
        readonly DBHelper _dbHelper = new DBHelper();
        readonly dealingClient _dealingClientData = new dealingClient();

        #region 申请人信息-基本信息
        public int InsertTCstmrApplicant(DataRow _dr, int _row,string commDB,SqlConnection _connection)
        {
            int nApplicantID = _dealingClientData.GetClientandApplicantIDByName(_dr["申请人代码"].ToString(), "Applicant",commDB, _connection);
            int nClientID = _dealingClientData.GetClientandApplicantIDByName(_dr["申请人代码"].ToString(), "Client",commDB, _connection);
            if (nClientID <= 0)
            {
                nClientID = -1;
            }

            string sName = _dr["中文名称、中译名"].ToString().Trim().Replace("'", "''");//中文名
            string sNativeName = _dr["外文名称"].ToString().Trim().Replace("'", "''");//英文名
            string sEmail = _dr["电子邮箱"].ToString().Trim().Replace("'", "''");
            string sFax = _dr["传真"].ToString().Trim().Replace("'", "''");
            string sWebsite = _dr["网址"].ToString().Trim().Replace("'", "''");
            string sMobile = _dr["手机"].ToString().Trim().Replace("'", "''");
            string sPhone = _dr["座机"].ToString().Trim().Replace("'", "''");
            string s_TrustDeedNo = _dr["总委托书"].ToString().Trim().Replace("'", "''");
            string sPayFeePerson = _dr["电子申请缴费人"].ToString().Trim().Replace("'", "''");
            string sAppType = _dr["申请人类型"].ToString().Trim().Replace("'", "''");
            string sIDNumber = _dr["身份证号或组织机构代码或统一社会信用代码"].ToString().Trim().Replace("'", "''");
            //费减资格备案年度	费减备案编号 
            string sFeeMitigationYear = _dr["费减资格备案年度"].ToString().Trim().Replace("'", "''");
            string sFeeMitigationNum = _dr["费减备案编号"].ToString().Trim().Replace("'", "''");
            //转换后数据
            int nCountry = _dealingClientData.GetAddressIDByName(_dr["国籍、注册国家（地区）"].ToString().Trim().Replace("'", "''"), commDB, _connection); 

            //string qualityRating = _dr["信用等级"] == null ? "" : _dr["信用等级"].ToString() == ""?"":"[信用等级:" + _dr["信用等级"].ToString() + "]\r\n";
            //string billingMode = _dr["账单方式"] == null ? "" : _dr["账单方式"].ToString() == "" ? "" : "[账单方式:" + _dr["账单方式"].ToString() + "]\r\n";
            //string currency = _dr["币种"] == null ? "" : _dr["币种"].ToString() == "" ? "" : "[币种:" + _dr["币种"].ToString() + "]";
            //string dunningCycle = _dr["账期（催款周期）"] == null ? "" : _dr["账期（催款周期）"].ToString() == "" ? "" : "[账期（催款周期）:" + _dr["账期（催款周期）"].ToString() + "]\r\n";
            //string language = _dr["通信语种"] == null ? "" : _dr["通信语种"].ToString() == "" ? "" : "[通信语种:" + _dr["通信语种"].ToString().Replace(",", "|") + "]\r\n";
            //string patentPerson = _dr["专利负责人"] == null ? "" : _dr["专利负责人"].ToString() == "" ? "" : "[专利负责人:" + _dr["专利负责人"].ToString() + "]\r\n";
            string sNotes = _dr["备注"].ToString().Trim().Replace("'", "''");

            string ctx = sNotes;//qualityRating + billingMode + currency + dunningCycle + language + patentPerson +

            if (!string.IsNullOrEmpty(sName) || !string.IsNullOrEmpty(sNativeName))
            {
                string strSql = "";
                if (nApplicantID > 0)
                {
                    strSql = "update    TCstmr_Applicant set s_Name='" + sName + "',s_NativeName='" + sNativeName + "',s_Mobile='" + sMobile + "',s_Phone='" + sPhone + "',s_Fax='" + sFax + "',s_Website='" + sWebsite + "',s_Email='" + sEmail + "',s_Notes='" + ctx.Replace("'", "''") + "'" +
                        ",s_AppType='" + sAppType + "',s_IDNumber='" + sIDNumber + "',n_Country=" + nCountry + ",n_ClientID=" + nClientID + ",s_FeeMitigationYear='" + sFeeMitigationYear + "',s_FeeMitigationNum='" + sFeeMitigationNum + "'";
                    if (sPayFeePerson.Equals("是") || sPayFeePerson.Equals("Y"))
                    {
                        strSql += ",s_PayFeePerson='Y'";
                    }
                    else
                    {
                        strSql += ",s_PayFeePerson='N'";
                    }
                    if (!string.IsNullOrEmpty(s_TrustDeedNo))
                    {
                        strSql += ",s_TrustDeedNo='" + s_TrustDeedNo + "',s_HasTrustDeed='Y',dt_EditDate='" + DateTime.Now + "'";
                    }
                    strSql += " where n_AppID=" + nApplicantID;
                }
                else
                {
                    strSql = "INSERT INTO dbo.TCstmr_Applicant(dt_CreateDate,dt_EditDate,dt_FirstCaseFromDate,dt_LastCaseFromDate,s_Creator,s_AppCode" +
                              ",s_Name,s_NativeName,s_Mobile,s_Phone,s_Fax,s_Website,s_Email,s_Notes,s_AppType,s_IDNumber,n_Country,n_ClientID,s_FeeMitigationYear,s_FeeMitigationNum";
                    string Values = "VALUES  ('" + DateTime.Now + "','" + DateTime.Now + "','" + DateTime.Now + "','" + DateTime.Now + "','administrator','" + _dr["申请人代码"].ToString() + "'" +
                                   ",'" + sName + "','" + sNativeName + "','" + sMobile + "','" + sPhone + "','" + sFax + "','" + sWebsite + "','" + sEmail + "','" + ctx.Replace("'", "''") + "','" + sAppType + "','" + sIDNumber + "'," + nCountry + "," + nClientID +
                                   ",'" + sFeeMitigationYear + "','" + sFeeMitigationNum + "'";

                    if (sPayFeePerson.Equals("是") || sPayFeePerson.Equals("Y"))
                    {
                        strSql += ",s_PayFeePerson";
                        Values += ",'Y'";
                    }
                    else
                    {
                        strSql += ",s_PayFeePerson";
                        Values += ",'N'";
                    }
                    if (!string.IsNullOrEmpty(s_TrustDeedNo))
                    {
                        strSql += ",s_TrustDeedNo,s_HasTrustDeed";
                        Values += ",'" + s_TrustDeedNo + "','Y'";
                    }
                    strSql = strSql + ")" + Values + ")";
                }
                return _dbHelper.InsertbySql(strSql, 0, commDB, _connection);
            }
            return 1;
        }
        #endregion

        #region 申请人-地址
        public int AddAppAddress(DataRow _dr, int _row,string commDB,SqlConnection _connection)
        {
            int nApplicantID = _dealingClientData.GetClientandApplicantIDByName(_dr["申请人代码"].ToString(), "Applicant",commDB, _connection);

            if (nApplicantID > 0)
            {
                return AddAPPRess(_dr, nApplicantID, "Applicant", commDB, _connection);
            }
            else
            {
                _dbHelper.InsertLog(0, _dr["申请人代码"].ToString(), _row, "申请人信息", "申请人-地址-" + _row, "未找申请人信息，申请人编号：" + _dr["申请人代码"].ToString(), "", commDB, _connection);
            }
            return 0;
        }
        #endregion

        #region 增加公共方法
        //地址
        private int AddAPPRess(DataRow _dr, int nAppID, string type, string commDB, SqlConnection _connection)
        {
            if (type.Equals("Applicant"))
            {
                string strSql = "delete TCstmr_AppAddress WHERE n_AppID=" + nAppID;
                _dbHelper.InsertbySql(strSql, 0, commDB, _connection);
            }
            int nCountry = _dealingClientData.GetAddressIDByName(_dr["居所或营业场所 （页签2：国家）1"].ToString().Trim().Replace("'", "''"), commDB, _connection);
            int nCountry2 = _dealingClientData.GetAddressIDByName(_dr["国家2"].ToString().Trim().Replace("'", "''"), commDB, _connection);
            int nCountry3 = _dealingClientData.GetAddressIDByName(_dr["国家3"].ToString().Trim().Replace("'", "''"), commDB, _connection);

            string sAddress = _dealingClientData.IPtype(_dr["地址1"].ToString().Trim().Replace("'", "''"));
            string sAddress2 = _dealingClientData.IPtype(_dr["地址2"].ToString().Trim().Replace("'", "''"));
            string sAddress3 = _dealingClientData.IPtype(_dr["地址3"].ToString().Trim().Replace("'", "''"));

            AddAPPAddress(nAppID, nCountry, _dr["省/州、自治区、直辖市 （中国客户）1"].ToString().Trim().Replace("'", "''"), _dr["市县（中国客户）1"].ToString().Trim().Replace("'", "''"), _dr["街道门牌（中国客户）1"].ToString().Trim().Replace("'", "''"), _dr["邮编（中国客户）1"].ToString().Trim().Replace("'", "''"), sAddress, commDB, _connection);
            AddAPPAddress(nAppID, nCountry2, _dr["省/州、自治区、直辖市 （中国客户）2"].ToString().Trim().Replace("'", "''"), _dr["市县（中国客户）2"].ToString().Trim().Replace("'", "''"), _dr["街道门牌2"].ToString().Trim().Replace("'", "''"), _dr["邮编（中国客户）2"].ToString().Trim().Replace("'", "''"), sAddress2, commDB, _connection);
            AddAPPAddress(nAppID, nCountry3, _dr["省/州、自治区、直辖市 （中国客户）3"].ToString().Trim().Replace("'", "''"), _dr["市县（中国客户）3"].ToString().Trim().Replace("'", "''"), _dr["街道门牌（中国客户）3"].ToString().Trim().Replace("'", "''"), _dr["邮编（中国客户）3"].ToString().Trim().Replace("'", "''"), sAddress3, commDB, _connection);

            return 1;
        }
        private void AddAPPAddress(int n_AppID, int nCountry, string sState, string sCity, string s_Street, string s_ZipCode, string sType, string commDB, SqlConnection _connection)
        {
            string strSql =
                "INSERT INTO dbo.TCstmr_AppAddress( n_AppID ,n_Country ,s_State ,s_City , s_Street ,s_ZipCode,s_Type)" +
                "VALUES(" + n_AppID + "," + nCountry + ",'" + sState + "','" + sCity + "','" + s_Street + "','" +
                s_ZipCode + "','" + sType + "')";

            _dbHelper.InsertbySql(strSql, 0, commDB, _connection);
        }

        #endregion

        #region 申请人-联系人
        public int AddAppContact(DataRow _dr, int _row, string commDB, SqlConnection _connection)
        { 
            if (!string.IsNullOrEmpty(_dr["申请人代码"].ToString()))
            {
                int nAppID = GetAppIDBysAppCode(_dr["申请人代码"].ToString(), commDB, _connection);
                int nCaseID = _dbHelper.GetIDbyName(_dr["案件编号"].ToString().Trim(), 2, _connection);

                if (nAppID > 0)
                {
                    int nLanguage = _dealingClientData.GetLanguageIDByName(_dr["语言"].ToString().Trim().Replace("'", "''"), commDB, _connection);
                    string sPhone = _dr["座机"].ToString().Trim().Replace("'", "''");
                    string sMobile = _dr["手机"].ToString().Trim().Replace("'", "''");
                    string sFirstName = _dr["名"].ToString().Trim().Replace("'", "''");
                    string sIPType = _dr["委托业务"].ToString().Trim().Replace("'", "''");
                    string sEmail = _dr["email"].ToString().Trim().Replace("'", "''");
                    string sDepartment = _dr["部门"].ToString().Trim().Replace("'", "''");
                    string sJobTitle = _dr["职位"].ToString().Trim().Replace("'", "''");
                    int nContactID = GetContactIDBysAppCode(sFirstName, sEmail, _dr["申请人代码"].ToString(), commDB, _connection);
                    string strSql = "";
                    if (nContactID > 0)
                    {
                        strSql = "update TCstmr_AppContact set s_Phone='" + sPhone + "',s_Mobile='" + sMobile + "',s_IPType='" + sIPType + "',n_Language=" + nLanguage + ",s_Department='" + sDepartment + "',s_JobTitle='" + sJobTitle + "' where n_AppID=" + nAppID;
                    }
                    else
                    {
                        strSql = " INSERT INTO dbo.TCstmr_AppContact( n_AppID ,s_FirstName , s_IPType ,n_Language ,s_Phone ,s_Mobile,s_Email,s_Department,s_JobTitle)" +
                            " VALUES  ( " + nAppID + ",'" + sFirstName + "','" + sIPType + "'," + nLanguage + ",'" + sPhone + "','" + sMobile + "','" + sEmail + "','" + sDepartment + "','" + sJobTitle + "')";
                    }
                    if (_dbHelper.InsertbySql(strSql, 0, commDB, _connection) > 0)
                    {
                        nContactID = GetContactIDBysAppCode(sFirstName, sEmail, _dr["申请人代码"].ToString(), commDB, _connection);
                        _dealingClientData.AddRess(_dr, nContactID, "AppContact", commDB, _connection);
                        if (nCaseID > 0)
                        {
                            return _dealingClientData.InsertTCaseContact(nCaseID, "Applicant", nContactID, "", _row, commDB, _connection); 
                        }
                    } 
                }
                else
                {
                    _dbHelper.InsertLog(nCaseID, _dr["案件编号"].ToString(), _row, "申请人信息", "申请人-联系人-" + _row, "未找申请人信息，申请人编号：" + _dr["申请人代码"].ToString(), "", commDB, _connection);
                }
            } 
            return 0;
        }
        private int GetAppIDBysAppCode(string sAppCode, string commDB, SqlConnection _connection)
        {
            string strSql = "SELECT n_AppID FROM TCstmr_Applicant WHERE s_AppCode='" + sAppCode + "'";
            return _dbHelper.InsertbySql(strSql, 0, commDB, _connection);
        }
        ////查找联系人是否存在
        private int GetContactIDBysAppCode(string sName, string sEmail, string sAppCode, string commDB, SqlConnection _connection)
        {
            string strSql = "SELECT n_ContactID FROM dbo.TCstmr_AppContact WHERE n_AppID in (SELECT n_AppID FROM TCstmr_Applicant WHERE s_AppCode='" + sAppCode + "') AND s_FirstName='" + sName.Trim() + "' AND s_Email='" + sEmail.Trim() + "'";
            return _dbHelper.InsertbySql(strSql, 0, commDB, _connection);
        }
        #endregion

        #region 申请人-财务
        public int AddAppBill(DataRow _dr,int _row,string form,string commDB,SqlConnection _connection)
        {
            //申请人代码	单位名称	纳税人识别号	注册地址	注册电话	开户行名称	银行账号	发票抬头
            int nAppID = _dealingClientData.GetClientandApplicantIDByName(_dr["代码"].ToString(), "Applicant", commDB, _connection);
            if (string.IsNullOrEmpty(_dr["代码"].ToString()))
            {
                nAppID = _dealingClientData.GetClientandApplicantIDByName(_dr["名称"].ToString(), "ApplicantName", commDB, _connection);
            }
            if (nAppID > 0)
            {
                string strSql = "update TCstmr_Applicant set s_AccountName='" + _dr["名称"].ToString() + "',s_AccountCode='" + _dr["开户行账号"].ToString() + "',s_TaxCode='" + _dr["纳税人识别号"].ToString() + "',s_RegAddress='" + _dr["注册地址"].ToString() + "',s_RegTel='" + _dr["注册电话"].ToString() + "',s_BankName='" + _dr["开户行名称"].ToString() + "',s_AccountNo='" + _dr["银行账户"].ToString() + "',s_InvoiceTo='" + _dr["发票抬头"].ToString() + "'" +
                                "  where n_AppID =" + nAppID;

                return _dbHelper.InsertbySql(strSql, 0, commDB, _connection);
            }
            else
            {
                if (string.IsNullOrEmpty(form))
                {
                    return _dealingClientData.AddBill(_dr, _row, "申请人", commDB, _connection);
                }
                else
                {
                    _dbHelper.InsertLog(0, _dr["代码"].ToString() + ":" + _dr["名称"].ToString(), _row, "申请人信息",
                                        "申请人-财务-" + _row,
                                        "未找申请人信息，申请人编号：" + _dr["代码"].ToString() + ":" + _dr["名称"].ToString(), "", commDB,
                                        _connection);
                }
            }
            return 0;
        }

        #endregion

        //总委托书号
        public int UpdateTotalCommissionNumber(DataRow row, int rowid, string commDB, SqlConnection _connection)
        {
            int result = 0;
            string sNo = row["CLIENT_NO"].ToString();

            //修改申请人为电子缴费单缴费人
            string strSql = " select n_AppID from TCstmr_Applicant  where s_AppCode='" + sNo + "'";
            int Num = _dbHelper.GetbySql(strSql, commDB, _connection);

            if (Num > 0)
            {
                strSql = " UPDATE TCstmr_Applicant  ";

                strSql += row["总委号1"].ToString() == "" ? "" : "set s_HasTrustDeed='Y',s_TrustDeedNo='" + row["总委号1"].ToString() + "'";
                if (strSql.Contains("set"))
                {
                    strSql += row["英文名称"].ToString() == "" ? "" : ",s_NativeName='" + row["英文名称"].ToString().Replace("'","''") + "'";
                }
                else
                {
                    strSql += row["英文名称"].ToString() == "" ? "" : " set s_NativeName='" + row["英文名称"].ToString().Replace("'","''") + "'";
                }
                if (strSql.Contains("set"))
                {
                    strSql += row["中文名称1"].ToString() == "" ? "" : ",s_Name='" + row["中文名称1"].ToString().Replace("'", "''") + "'";
                }
                else
                {
                    strSql += row["中文名称1"].ToString() == "" ? "" : " set s_Name='" + row["中文名称1"].ToString().Replace("'", "''") + "'";
                }
                strSql += "  WHERE n_AppID=" + Num;
                strSql += "   update TCase_Applicant set  s_TrustDeedNo='" + row["总委号1"].ToString() + "' where n_ApplicantID=" + Num;
                
                //增加译名
                if (!string.IsNullOrEmpty(row["中文名称2"].ToString()) || !string.IsNullOrEmpty(row["总委号2"].ToString()))
                {
                    string sql = " SELECT n_ID FROM  dbo.TCstmr_AppTransLatedName   WHERE n_AppID=" + Num +
                                 " AND s_AppTransLatedName='" + row["中文名称2"].ToString() + "' AND s_TrustdeedNum='" +
                                 row["总委号2"].ToString() + "' ";
                    if (_dbHelper.GetbySql(sql, commDB, _connection) <= 0)
                    {
                        strSql +=
                            " INSERT INTO dbo.TCstmr_AppTransLatedName( s_AppTransLatedName ,s_TrustdeedNum ,s_AppTransLatedNameUse ,n_AppID)VALUES  ('" +
                            row["中文名称2"].ToString() + "','" + row["总委号2"].ToString() + "','P, T, C, D, O'," + Num + ")";
                    }
                }
                if (!string.IsNullOrEmpty(row["中文名称3"].ToString()) || !string.IsNullOrEmpty(row["总委号3"].ToString()))
                {
                    string sql = " SELECT n_ID FROM  dbo.TCstmr_AppTransLatedName   WHERE n_AppID=" + Num +
                                 " AND s_AppTransLatedName='" + row["中文名称3"].ToString() + "' AND s_TrustdeedNum='" +
                                 row["总委号3"].ToString() + "' ";
                    if (_dbHelper.GetbySql(sql, commDB, _connection) <= 0)
                    {
                        strSql +=
                            " INSERT INTO dbo.TCstmr_AppTransLatedName( s_AppTransLatedName ,s_TrustdeedNum ,s_AppTransLatedNameUse ,n_AppID)VALUES  ('" +
                            row["中文名称3"].ToString() + "','" + row["总委号3"].ToString() + "','P, T, C, D, O'," + Num + ")";
                    }
                }
                result = _dbHelper.InsertbySql(strSql, 0, commDB, _connection); 
                
                if(result<=0)
                {
                    _dbHelper.InsertLog(Num, sNo, rowid, "总委托书号", "总委托书号-" + rowid, "未找申请人信息，申请人编号：" + sNo, strSql.Replace("'","''"), commDB, _connection);
                }
            }
            else
            {
                //_dbHelper.InsertLog(Num, sNo, rowid, "总委托书号", "总委托书号-" + rowid, "未找申请人信息，申请人编号：" + sNo, "", commDB, _connection);
                strSql = "INSERT INTO dbo.TCstmr_Applicant(dt_CreateDate,dt_EditDate,dt_FirstCaseFromDate,dt_LastCaseFromDate,s_Creator,s_AppCode,s_Name,s_NativeName)";
                string Values = "VALUES  ('" + DateTime.Now + "','" + DateTime.Now + "','" + DateTime.Now + "','" + DateTime.Now + "','administrator','" + row["CLIENT_NO"].ToString() + "'" +
                               ",'" + row["中文名称1"].ToString() + "','" + row["英文名称"].ToString() + "')";
                if (_dbHelper.InsertbySql(strSql + Values, 0, commDB, _connection) > 0)
               {
                   UpdateTotalCommissionNumber(row, rowid, commDB, _connection);
               }  
            }
            return result;
        }
    }
}
