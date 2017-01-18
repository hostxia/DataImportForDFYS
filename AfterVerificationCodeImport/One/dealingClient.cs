using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace AfterVerificationCodeImport
{
    class dealingClient
    {
        readonly DBHelper _dbHelper=new DBHelper();
      
        #region 联系人
        public int AddClientContact(DataRow _dr, int _row, string commDB, SqlConnection _conn)
        {
            int nClientID = GetClientIDByCaseSerialandCode(_dr["案件编号"].ToString(), "Client", _dr["客户代码"].ToString(), commDB, _conn);
            int nCaseID = _dbHelper.GetIDbyName(_dr["案件编号"].ToString().Trim(), 2, _conn);
            string strSql = "";
            if (nClientID > 0)
            {
                int nLanguage = GetLanguageIDByName(_dr["语言"].ToString().Trim().Replace("'", "''"), commDB, _conn);
                string sPhone = _dr["座机"].ToString().Trim().Replace("'", "''");
                string sMobile = _dr["手机"].ToString().Trim().Replace("'", "''");
                string sFirstName = _dr["名"].ToString().Trim().Replace("'", "''");
                string sIPType = _dr["委托业务"].ToString().Trim().Replace("'", "''");
                string sEmail = _dr["email"].ToString().Trim().Replace("'", "''");
                int nContactID = GetClientIDByCaseSerial(sFirstName, sEmail, nClientID, "Client", commDB, _conn);
                if (nContactID > 0)
                {
                    strSql = "update TCstmr_ClientContact set s_Phone='" + sPhone + "',s_Mobile='" + sMobile + "',s_IPType='" + sIPType + "',n_Language=" + nLanguage + " where n_ContactID=" + nContactID;
                }
                else
                {
                    strSql = " INSERT INTO dbo.TCstmr_ClientContact( n_ClientID ,s_FirstName , s_IPType ,n_Language ,s_Phone ,s_Mobile,s_Email)" +
                        " VALUES  ( " + nClientID + ",'" + sFirstName + "','" + sIPType + "'," + nLanguage + ",'" + sPhone + "','" + sMobile + "','" + sEmail + "')";

                }
                if (_dbHelper.InsertbySql(strSql, _row,commDB,_conn) > 0)
                {
                    nContactID = GetClientIDByCaseSerial(sFirstName, sEmail, nClientID, "Client", commDB, _conn);
                    AddRess(_dr, nContactID, "Contact",commDB,_conn);
                    if (nCaseID > 0)
                    {
                        InsertTCaseContact(nCaseID, "Client", nContactID, "",_row,commDB, _conn);
                    }
                }
                return 1;
            }
            else
            {
                _dbHelper.InsertLog(nCaseID, _dr["案件编号"].ToString(), _row, "客户信息","客户-联系人-"+_row,"未找案件下找到客户信息，客户编号：" + _dr["客户代码"].ToString(), strSql,commDB, _conn);
            }
            //信函抬头	账单抬头	   
            return 0;
        }

        //增加案件与联系人关系
        public int InsertTCaseContact(int nCaseID, string sContactType, int nContactID, string sIdentity, int _rowID, string commDB, SqlConnection _conn)
        {
            string strSql = "SELECT TOP 1 n_Sequence FROM  dbo.TCase_Contact WHERE n_CaseID=" + nCaseID + " ORDER BY n_Sequence DESC ";
            int nSequence = _dbHelper.GetbySql(strSql, commDB, _conn) + 1;
            strSql = "SELECT COUNT(*) AS sumcount FROM  dbo.TCase_Contact WHERE n_CaseID=" + nCaseID + " AND s_ContactType='" + sContactType + "' AND n_ContactID=" + nContactID;
            if (_dbHelper.GetbySql(strSql, commDB, _conn) <= 0)
            {
                strSql = "INSERT INTO dbo.TCase_Contact( n_CaseID ,s_ContactType , n_ContactID ,n_Sequence ,s_Identity)" +
                        "VALUES  (" + nCaseID + ",'" + sContactType + "'," + nContactID + "," + nSequence + ",'" + sIdentity + "')";
               return _dbHelper.InsertbySql(strSql, _rowID,commDB,_conn);
            }
            return 0;
        }

        //根据案件查询客户ID
        public int GetClientIDByCaseSerialandCode(string sCaseSerial, string type, string sClientCode, string commDB, SqlConnection conn)
        {
            string strSql = " SELECT a.n_ClientID FROM dbo.TCase_Base  a   LEFT JOIN dbo.TCstmr_Client b ON a.n_ClientID = b.n_ClientID  WHERE s_CaseSerial ='" + sCaseSerial + "' AND s_ClientCode='" + sClientCode + "'";
            if (type.Equals("Agency"))
            {
                strSql = "  SELECT n_CoopAgencyToID FROM dbo.TCase_Base  a   LEFT JOIN dbo.TCstmr_CoopAgency b ON a.n_CoopAgencyToID = b.n_AgencyID    WHERE s_CaseSerial='" + sCaseSerial + "' AND s_Code='" + sClientCode + "'";
            }
            return _dbHelper.GetbySql(strSql, commDB, conn);
        }

        //查找联系人是否存在
        public int GetClientIDByCaseSerial(string sName, string sEmail, int nClientID, string type, string commDB, SqlConnection conn)
        {
            string strSql = "  SELECT n_ContactID FROM dbo.TCstmr_ClientContact WHERE n_ClientID=" + nClientID + " AND s_FirstName='" + sName.Trim() + "' AND s_Email='" + sEmail.Trim() + "'";
            if (type.Equals("Agency"))
            {
                strSql = "  SELECT n_ContactID FROM dbo.TCstmr_AgencyContact WHERE n_AgencyID=" + nClientID + " AND s_FirstName='" + sName.Trim() + "' AND s_Email='" + sEmail.Trim() + "'";
            }
            return _dbHelper.GetbySql(strSql, commDB, conn);
        }
        #endregion

        #region 客户信息-基本信息
        public int InsertTCstmrClient(DataRow _dr, int _row, string commDB, SqlConnection _connection)
        {
            int nClientID = GetClientandApplicantIDByName(_dr["代码"].ToString(), "Client",commDB,_connection);
            string strSql = "";

            string sName = _dr["中文名称"].ToString().Trim().Replace("'", "''");//中文名
            string sNativeName = _dr["英文名称"].ToString().Trim().Replace("'", "''");//英文名
            string sEmail = _dr["电子邮箱"].ToString().Trim().Replace("'", "''");
            string sFax = _dr["传真"].ToString().Trim().Replace("'", "''");
            string sWebsite = _dr["网址"].ToString().Trim().Replace("'", "''");
            string sMobile = _dr["手机"].ToString().Trim().Replace("'", "''");
            string sPhone = _dr["座机"].ToString().Trim().Replace("'", "''");
            string sNotes = _dr["备注"].ToString().Trim().Replace("'", "''");
            string sType = _dr["客户类别（账单方式）"].ToString().Trim().Replace("'", "''");
            string sCredit = _dr["信用等级"].ToString().Trim().Replace("'", "''");
            string sState = _dr["省、州"].ToString().Trim().Replace("'", "''");
            string sCity = _dr["市县"].ToString().Trim().Replace("'", "''");

            //转换后数据
            int nCountry = GetAddressIDByName(_dr["国家"].ToString().Trim().Replace("'", "''"), commDB, _connection);
            string sIPType = IPtype(_dr["委托业务"].ToString().Trim().Replace("'", "''"));//P:专利；T:商标；D：域名；C：版权 O：其它
            int nPatentChargerID = GetEmployeeIDByName(_dr["专利负责人"].ToString().Trim().Replace("'", "''"), commDB, _connection);

            string[] d = _dr["通信语种"].ToString().Trim().Replace("'", "''").Split(',');
            int nLanguage = 0;
            if (d.Length > 0 && !string.IsNullOrEmpty(d[0]))
            {
                nLanguage = GetLanguageIDByName(d[0], commDB, _connection);
            }
            //string dunningCycle = _dr["账期（催款周期）"] == null ? "" :_dr["账期（催款周期）"].ToString()==""?"":"[账期（催款周期）:" + _dr["账期（催款周期）"].ToString() + "]\r\n";
            string slanguage = _dr["通信语种"] == null ? "" : _dr["通信语种"].ToString() == "" ? "" : "[通信语种:" + _dr["通信语种"].ToString().Replace(",", "|") + "]\r\n";

            sNotes = slanguage + sNotes;

            if (nClientID > 0)
            {
                strSql = "update TCstmr_Client set s_Name='" + sName + "',s_NativeName='" + sNativeName + "',s_Email='" + sEmail + "',s_Fax='" + sFax + "',s_Website='" + sWebsite + "',s_Mobile='" + sMobile + "',s_Phone='" + sPhone + "',s_Notes='" + sNotes + "',s_Type='" + sType + "',s_Credit='" + sCredit + "'" +
                         ",n_Country=" + nCountry + ",s_State='" + sState + "',s_City='" + sCity + "',n_Language=" + nLanguage + ",n_PatentChargerID=" + nPatentChargerID + ",s_IPType='" + sIPType + "',dt_EditDate='" + DateTime.Now + "' where n_ClientID=" + nClientID + "";
                //InsertClientFeePolicy(_dr["币种"].ToString(), nClientID, _row,commDB,_connection);
                return _dbHelper.InsertbySql(strSql, _row,commDB,_connection);
            }
            else
            {
                if (!string.IsNullOrEmpty(sName) || !string.IsNullOrEmpty(sNativeName) || !string.IsNullOrEmpty(_dr["代码"].ToString()))
                {
                    strSql =
                        "INSERT INTO dbo.TCstmr_Client(dt_CreateDate,dt_EditDate,s_Creator,s_ClientCode,s_Name,s_NativeName,s_Email,s_Fax,s_Website,s_Mobile,s_Phone,s_Notes,s_Type,s_Credit,n_Country,s_State,s_City,n_Language,n_PatentChargerID,s_IPType) " +
                        "VALUES('" + DateTime.Now + "','" + DateTime.Now + "','administrator','" + _dr["代码"].ToString() +
                        "','" + sName + "','" + sNativeName + "','" + sEmail + "','" + sFax + "','" + sWebsite + "','" +
                        sMobile + "','" + sPhone + "','" + sNotes + "','" + sType + "','" + sCredit + "'" +
                        "," + nCountry + ",'" + sState + "','" + sCity + "'," + nLanguage + "," + nPatentChargerID +
                        ",'" + sIPType + "')";
                    if (_dbHelper.InsertbySql(strSql, _row,commDB, _connection) > 0)
                    {
                        nClientID = GetClientandApplicantIDByName(_dr["代码"].ToString(), "Client",commDB, _connection);
                        //InsertClientFeePolicy(_dr["币种"].ToString(), nClientID, _row,commDB, _connection);
                    }
                    return 1;
                }
            }
            return 0;
        }

        //增加币种
        //private void InsertClientFeePolicy(string currency, int nClientID, int _row,string commDB, SqlConnection _connection)
        //{
        //    if (!string.IsNullOrEmpty(currency))
        //    {
        //        string strSql = "";
        //        string[] arr = currency.Split(',');
        //        foreach (var item in arr)
        //        {
        //            string selectSql = "SELECT n_ID FROM dbo.TCode_Currency WHERE s_Name='" + item.Trim() +
        //                               "' OR s_CurrencyCode='" + item.Trim() + "'";
        //            int current = _dbHelper.GetbySql(selectSql, commDB, _connection);
        //            if (current > 0)
        //            {
        //                strSql =
        //                    "INSERT INTO dbo.TCstmr_ClientFeePolicy( n_ClientID , n_ChargeCurrency , s_IPType ,dt_EditDate , s_BusinessType ,s_PTCType)" +
        //                    "VALUES (" + nClientID + "," + current + ",'P','" + DateTime.Now + "','-1','-1')   ";
        //            }
        //            else
        //            {
        //                _dbHelper.InsertLog(0, "", _row, "客户信息", "客户-" + _row, "已存在当前币种，无需添加,币种类型:" + item.Trim(),
        //                                    strSql, commDB, _connection);
        //            }
        //        }
        //        _dbHelper.InsertbySql(strSql, _row, commDB, _connection);
        //    }
        //} 

        //查找地址信息
        public int GetAddressIDByName(string countryName, string commDB, SqlConnection _connection)
        {
            if (!string.IsNullOrEmpty(countryName))
            {
                string strSql = "SELECT n_ID FROM dbo.TCode_Country  WHERE  s_Name='" + countryName.Trim() + "' OR s_CountryCode='" + countryName.Trim() + "' OR s_OtherName='" + countryName.Trim() + "'";
                if (_dbHelper.GetbySql(strSql, commDB, _connection) > 0)
                {
                    return _dbHelper.GetbySql(strSql, commDB, _connection);
                }
            }
            return -1;
        }

        //查找语言ID
        public int GetLanguageIDByName(string sLanguageName, string commDB, SqlConnection _connection)
        {
            if (!string.IsNullOrEmpty(sLanguageName))
            {
                string strSql = "SELECT n_ID FROM dbo.TCode_Language  WHERE  s_Name='" + sLanguageName.Trim() + "' or s_LanguageCode='" + sLanguageName.Trim() + "' or s_OtherName='" + sLanguageName.Trim() + "'";
                return _dbHelper.GetbySql(strSql, commDB, _connection); 
            }
            return 0;
        }

        //专利负责人
        private int GetEmployeeIDByName(string sEmployeeName,string commDB, SqlConnection _connection)
        {
            if (!string.IsNullOrEmpty(sEmployeeName))
            {
                string strSql = "SELECT n_ID FROM dbo.TCode_Employee  WHERE  s_Name='" + sEmployeeName.Trim() + "' or s_InternalCode='" + sEmployeeName.Trim() + "'";
                return _dbHelper.GetbySql(strSql, commDB, _connection); 
            }
            return 0;
        }
        #endregion

        #region 客户-地址
        public int AddClientAddress(DataRow _dr, int _row, string commDB, SqlConnection _connection)
        {
            int nClientID = GetClientandApplicantIDByName(_dr["代码"].ToString(), "Client",commDB, _connection);
            if (nClientID > 0)
            {
                return AddRess(_dr, nClientID, "Client", commDB,_connection);
            }
            else
            {
                _dbHelper.InsertLog(0, "", _row, "客户信息", "客户-地址" + _row, "未找到当前客户信息:" + _dr["代码"].ToString(), "", commDB,_connection);
            }
            return 0;
        }
        #endregion

        #region 增加公共方法
        //是否存在此地址
        public void GetIDbyClientAddress(string type, int nContactID, string commDB, SqlConnection _connection)
        {
            string strSql = "delete TCstmr_ClientAddress WHERE n_ClientID=" + nContactID;
            if (type.Equals("Contact"))
            {
                strSql = "delete TCstmr_ClientConAddress WHERE n_ContactID=" + nContactID;
            }
            else if (type.Equals("AppContact"))
            {
                strSql = "delete TCstmr_AppConAddress WHERE n_ContactID=" + nContactID;
            }
            else if (type.Equals("CoopAgency"))
            {
                strSql = "delete TCstmr_AgencyAddress WHERE n_AgencyID=" + nContactID;
            }
            else if (type.Equals("Agency"))
            {
                strSql = "delete TCstmr_AgencyConAddress WHERE n_ContactID=" + nContactID;
            }
            _dbHelper.InsertbySql(strSql, 0,commDB,_connection);
        }

        private void AddClientAddress(int n_ClientID, int nCountry, string sState, string sCity, string s_Street, string s_ZipCode, string type, string saddress, int nSequence, string sTitleAddress,string commDB,SqlConnection _connection)
        {
            sState = sState.Replace(",", ",,");
            sCity = sCity.Replace(",", ",,");
            s_Street = s_Street.Replace(",", ",,");
            string strSql = "INSERT INTO dbo.TCstmr_ClientAddress( n_ClientID ,n_Country ,s_State ,s_City , s_Street ,s_ZipCode,s_Type,s_TitleAddress)" +
                           "VALUES(" + n_ClientID + "," + nCountry + ",'" + sState + "','" + sCity + "','" + s_Street + "','" + s_ZipCode + "','" + saddress + "','" + sTitleAddress + "')";
            if (type.Equals("Contact"))
            {
                strSql = "INSERT INTO dbo.TCstmr_ClientConAddress( n_ContactID ,n_Country ,s_State ,s_City , s_Street ,s_ZipCode,s_Type,s_TitleAddress)" +
                           "VALUES(" + n_ClientID + "," + nCountry + ",'" + sState + "','" + sCity + "','" + s_Street + "','" + s_ZipCode + "','" + saddress + "','" + sTitleAddress + "')";
            }
            else if (type.Equals("AppContact"))
            {
                strSql = "INSERT INTO dbo.TCstmr_AppConAddress( n_ContactID ,n_Country ,s_State ,s_City , s_Street ,s_ZipCode,s_Type,n_Sequence,s_TitleAddress)" +
                               "VALUES(" + n_ClientID + "," + nCountry + ",'" + sState + "','" + sCity + "','" + s_Street + "','" + s_ZipCode + "','" + saddress + "'," + nSequence + ",'" + sTitleAddress + "')";
            }
            else if (type.Equals("Agency"))
            {
                strSql = "INSERT INTO dbo.TCstmr_AgencyConAddress( n_ContactID ,n_Country ,s_State ,s_City , s_Street ,s_ZipCode,s_Type)" +
                            "VALUES(" + n_ClientID + "," + nCountry + ",'" + sState + "','" + sCity + "','" + s_Street + "','" + s_ZipCode + "','" + saddress + "')";
            }
            _dbHelper.InsertbySql(strSql, 0,commDB, _connection);
        }

        //地址
        public int AddRess(DataRow _dr, int nClientID, string type, string commDB, SqlConnection _connection) 
        {
            GetIDbyClientAddress(type, nClientID, commDB,_connection);
            int nCountry = GetAddressIDByName(_dr["国家1"].ToString().Trim().Replace("'", "''"), commDB, _connection);
            int nCountry2 = GetAddressIDByName(_dr["国家2"].ToString().Trim().Replace("'", "''"), commDB, _connection);
            int nCountry3 = GetAddressIDByName(_dr["国家3"].ToString().Trim().Replace("'", "''"), commDB, _connection);

            string sAddress = IPtype(_dr["地址1"].ToString().Trim().Replace("'", "''"));
            string sAddress2 = IPtype(_dr["地址2"].ToString().Trim().Replace("'", "''"));
            string sAddress3 = IPtype(_dr["地址3"].ToString().Trim().Replace("'", "''"));

            AddClientAddress(nClientID, nCountry, _dr["省、州1"].ToString().Trim().Replace("'", "''"), _dr["市县1"].ToString().Trim().Replace("'", "''"), _dr["街道门牌1"].ToString().Trim().Replace("'", "''"), _dr["邮政编码1"].ToString().Trim().Replace("'", "''"), type, sAddress, 0,"" ,commDB, _connection);
            AddClientAddress(nClientID, nCountry2, _dr["省、州2"].ToString().Trim().Replace("'", "''"), _dr["市县2"].ToString().Trim().Replace("'", "''"), _dr["街道门牌2"].ToString().Trim().Replace("'", "''"), _dr["邮政编码2"].ToString().Trim().Replace("'", "''"), type, sAddress2, 1,"",commDB, _connection);
            AddClientAddress(nClientID, nCountry3, _dr["省、州3"].ToString().Trim().Replace("'", "''"), _dr["市县3"].ToString().Trim().Replace("'", "''"), _dr["街道门牌3"].ToString().Trim().Replace("'", "''"), _dr["邮政编码3"].ToString().Trim().Replace("'", "''"), type, sAddress3, 2,"",commDB, _connection);

            return 1;
        }

        //类型转换
        public string IPtype(string _content)
        {
            //P:专利；T:商标；D：域名；C：版权 O：其它B, M, O, E
            string strHtml = "";
            string[] arry = _content.Replace('，', ',').Split(',');
            foreach (string t in arry.Where(t => !string.IsNullOrEmpty(t)))
            {
                if (t.Equals("专利"))
                {
                    strHtml += "P,";
                }
                else if (t.Equals("商标"))
                {
                    strHtml += "T,";
                }
                else if (t.Equals("域名"))
                {
                    strHtml += "D,";
                }
                else if (t.Equals("版权"))
                {
                    strHtml += "C,";
                }
                else if (t.Equals("法律&其他案"))
                {
                    strHtml += "O,";
                }//账单 B  转函地址 M 办公地址 O  办公地址（外文）E   
                else if (t.Equals("账单地址"))
                {
                    strHtml += "B,";
                }
                else if (t.Equals("转函地址"))
                {
                    strHtml += "M,";
                }
                else if (t.Equals("办公地址"))
                {
                    strHtml += "O,";
                }
                else if (t.Equals("办公地址(外文)"))
                {
                    strHtml += "E,";
                }
            }
            if (strHtml.Length > 0)
            {
                strHtml = strHtml.Substring(0, strHtml.Length - 1);
            }
            return strHtml;
        }
        #endregion

        #region 客户-财务
        public int AddBill(DataRow _dr,int _row, string form, string commDB, SqlConnection _connection)
        {
            int nClientID = GetClientandApplicantIDByName(_dr["代码"].ToString(), "Client",commDB, _connection);
            if (string.IsNullOrEmpty(_dr["代码"].ToString()))
            {
                nClientID = GetClientandApplicantIDByName(_dr["名称"].ToString(), "ClientName", commDB, _connection);
            }
            if (nClientID > 0)
            {
                string strSql = "update TCstmr_Client set s_AccountName='" + _dr["名称"].ToString() + "',s_TaxCode='" + _dr["纳税人识别号"].ToString() + "',s_RegAddress='" + _dr["注册地址"].ToString() + "'," +
                                "s_RegTel='" + _dr["注册电话"].ToString() + "',s_BankName='" + _dr["开户行名称"].ToString() + "',s_AccountNo='" + _dr["银行账户"].ToString() + "',s_AccountCode='" + _dr["开户行账号"].ToString() + "'" +
                                ",s_InvoiceTo='" + _dr["发票抬头"].ToString() + "'  where n_ClientID =" + nClientID;
                return _dbHelper.InsertbySql(strSql, 0, commDB, _connection);
            }
            else
            {
                if (string.IsNullOrEmpty(form))
                {
                    var _dealingApplicant = new dealingApplicant();
                    return _dealingApplicant.AddAppBill(_dr, _row, "客户", commDB, _connection);
                }
                else
                {
                    _dbHelper.InsertLog(0, _dr["代码"].ToString() + ":" + _dr["名称"].ToString(), _row, "客户信息",
                                        "客户-财务" + _row, "未找到当前客户信息:" + _dr["代码"].ToString() + ":" + _dr["名称"].ToString(),
                                        "", commDB, _connection);
                }
            }

            return 0;
        }
        //客户ID
        public int GetClientandApplicantIDByName(string sCode, string type, string commDB, SqlConnection _connection)
        {
            string strSql = "";
            if (!string.IsNullOrEmpty(sCode) && type.Equals("Client"))
            {
                strSql = "select  n_ClientID from TCstmr_Client where s_ClientCode='" + sCode.Trim() + "'";
                return _dbHelper.GetbySql(strSql, commDB, _connection);
            }
            else if (!string.IsNullOrEmpty(sCode) && type.Equals("ClientName"))
            {
                strSql = "select  n_ClientID from TCstmr_Client where s_Name='" + sCode.Trim() + "' or s_NativeName='" + sCode + "'";
                return _dbHelper.GetbySql(strSql, commDB, _connection);
            }
            else if (!string.IsNullOrEmpty(sCode) && type.Equals("Applicant"))
            {
                strSql = "select  n_AppID from TCstmr_Applicant where s_AppCode='" + sCode.Trim() + "'";
                return _dbHelper.GetbySql(strSql, commDB, _connection);
            }
            else if (!string.IsNullOrEmpty(sCode) && type.Equals("ApplicantName"))
            {
                strSql = "select  n_AppID from TCstmr_Applicant where s_Name='" + sCode.Trim() + "' or s_NativeName='" + sCode + "'";
                return _dbHelper.GetbySql(strSql, commDB, _connection);
            }
            else if (!string.IsNullOrEmpty(sCode) && type.Equals("CoopAgency"))
            {
                strSql = "select  n_AgencyID from TCstmr_CoopAgency WHERE s_Code='" + sCode.Trim() + "'";
                return _dbHelper.GetbySql(strSql, commDB, _connection);
            }
            else if (!string.IsNullOrEmpty(sCode) && type.Equals("CoopAgencyName"))
            {
                strSql = "select  n_AgencyID from TCstmr_CoopAgency WHERE s_Name='" + sCode.Trim() + "' or s_NativeName='" + sCode + "'";
                return _dbHelper.GetbySql(strSql, commDB, _connection);
            }
            return 0;
        }
        #endregion

        #region 更新客户和申请人关系
        public void updateApplicantandClient(string commDB, SqlConnection _connection)
        {
            const string strSql = " UPDATE TCstmr_Applicant SET n_ClientID=(SELECT n_ClientID FROM TCstmr_Client WHERE s_ClientCode=s_AppCode)  WHERE  dbo.TCstmr_Applicant.n_ClientID=-1 AND s_AppCode IS NOT NULL AND s_AppCode!='' AND s_AppCode IN (SELECT s_AppCode FROM dbo.TCstmr_Client WHERE s_ClientCode=s_AppCode) " +
                                  " UPDATE TCstmr_Client SET n_ApplicantID=(SELECT n_AppID FROM TCstmr_Applicant WHERE s_AppCode=TCstmr_Client.s_ClientCode)  WHERE  n_ApplicantID=-1 AND s_ClientCode IS NOT NULL AND s_ClientCode!='' AND s_ClientCode IN (SELECT s_AppCode FROM TCstmr_Applicant WHERE s_AppCode=TCstmr_Client.s_ClientCode) ";
            _dbHelper.InsertbySql(strSql, 0,commDB, _connection);
        }

        #endregion
    }
}
