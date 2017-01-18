using System;
using System.Data;
using System.Data.SqlClient;

namespace AfterVerificationCodeImport
{
    class dealingAgency
    {
        readonly DBHelper _dbHelper = new DBHelper();
        readonly dealingClient _dealingClientData = new dealingClient();


        #region 外代理信息-基本信息
        public int InsertTCstmrCoopAgency(DataRow _dr, int _row,string commDB, SqlConnection _connection)
        {
            int nAgencyID = _dealingClientData.GetClientandApplicantIDByName(_dr["编码"].ToString(), "CoopAgency", commDB, _connection);
            int nClientID = _dealingClientData.GetClientandApplicantIDByName(_dr["编码"].ToString(), "Client", commDB, _connection);
            if (nClientID <= 0)
            {
                nClientID = -1;
            }
            string strSql = "";
            string sName = _dr["中文名称、中译名"].ToString().Trim().Replace("'", "''");//中文名
            string sNativeName = _dr["名称"].ToString().Trim().Replace("'", "''");//英文名
            int nLagnguage = _dealingClientData.GetLanguageIDByName(_dr["语种"].ToString().Trim().Replace("'", "''"), commDB, _connection);
            int nPayCurrency = GetCurrencyIDByName(_dr["结算币种"].ToString().Trim().Replace("'", "''"), commDB, _connection);
            string sIPType = _dealingClientData.IPtype(_dr["委托业务"].ToString().Trim().Replace("'", "''"));//P:专利；T:商标；D：域名；C：版权 O：其它
            string sCredit = _dr["信用等级"].ToString().Trim().Replace("'", "''");

            string sEmail = _dr["电子邮箱"].ToString().Trim().Replace("'", "''");
            string sFax = _dr["传真"].ToString().Trim().Replace("'", "''");
            string sWebsite = _dr["网址"].ToString().Trim().Replace("'", "''");
            string sMobile = _dr["手机"].ToString().Trim().Replace("'", "''");
            string sPhone = _dr["座机"].ToString().Trim().Replace("'", "''");
            string sNotes = _dr["备注"].ToString().Trim().Replace("'", "''");

            string sCountry = _dr["可代理国家（地区）"] != null ? _dr["可代理国家（地区）"].ToString() == "" ? "" : "[可代理国家（地区）:" + _dr["可代理国家（地区）"].ToString().Trim().Replace("'", "''").Replace(",", "|") + "]\r\n" : "";
            //string billingMode = _dr["账单方式"] == null ? "" : _dr["账单方式"].ToString() == "" ? "" : "[账单方式:" + _dr["账单方式"].ToString() + "]\r\n";
            //string dunningCycle = _dr["账期（催款周期）"] == null ? "" : _dr["账期（催款周期）"].ToString() == "" ? "" : "[账期（催款周期）:" + _dr["账期（催款周期）"].ToString() + "]\r\n";
            //string patentPerson = _dr["专利负责人"] == null ? "" : _dr["专利负责人"].ToString() == "" ? "" : "[专利负责人:" + _dr["专利负责人"].ToString() + "]\r\n";

            sNotes = sCountry + sNotes;//billingMode + dunningCycle + patentPerson +
            if (nAgencyID > 0)
            {
                strSql = "update   TCstmr_CoopAgency set s_Name='" + sName + "',s_NativeName='" + sNativeName + "',s_Mobile='" + sMobile + "',s_Phone='" + sPhone + "',s_Fax='" + sFax + "',s_Website='" + sWebsite + "',s_Email='" + sEmail + "',s_Notes='" + sNotes + "'" +
                    ",n_Language=" + nLagnguage + ",n_PayCurrency=" + nPayCurrency + ",s_IPType='" + sIPType + "',s_Credit='" + sCredit + "' where n_AgencyID=" + nAgencyID;
            }
            else
            {
                if (!string.IsNullOrEmpty(sName) || !string.IsNullOrEmpty(sNativeName) || !string.IsNullOrEmpty(_dr["编码"].ToString()))
                {
                    strSql =
                        "INSERT INTO dbo.TCstmr_CoopAgency(dt_CreateDate,dt_EditDate,dt_FirstCaseFromDate,dt_LastCaseFromDate,s_Creator,s_Code" +
                        ",s_Name,s_NativeName,s_Mobile,s_Phone,s_Fax,s_Website,s_Email,s_Notes,n_Language,n_PayCurrency,s_IPType,s_Credit,n_ClientID)" +
                        "VALUES  ('" + DateTime.Now + "','" + DateTime.Now + "','" + DateTime.Now + "','" + DateTime.Now +
                        "','administrator','" + _dr["编码"].ToString() + "'" +
                        ",'" + sName + "','" + sNativeName + "','" + sMobile + "','" + sPhone + "','" + sFax + "','" +
                        sWebsite + "','" + sEmail + "','" + sNotes + "'" +
                        "," + nLagnguage + "," + nPayCurrency + ",'" + sIPType + "','" + sCredit + "'," + nClientID +
                        ")";
                }
            }
            return !string.IsNullOrEmpty(strSql) ? _dbHelper.InsertbySql(strSql, _row, commDB, _connection) : 0;
        }

        #endregion

        #region 外代理信息-地址
        public int AddAgencyAddress(DataRow _dr, int _row, string commDB, SqlConnection _connection)
        {
            int nAgencyID = _dealingClientData.GetClientandApplicantIDByName(_dr["编码"].ToString(), "CoopAgency", commDB, _connection);

            if (nAgencyID > 0)
            {
                AddAgencyRess(_dr, nAgencyID, "CoopAgency", commDB, _connection); 
            }
            else
            {
                _dbHelper.InsertLog(0, _dr["编码"].ToString(), _row, "外代理", "外代理-地址-" + _row, "未找外代理信息,编号：" + _dr["编码"].ToString(), "", commDB, _connection); 
            }
            return 1;
        }
        #endregion

        #region 增加公共方法
        //币种
        private int GetCurrencyIDByName(string CurrencyName, string commDB, SqlConnection _connection)
        {
            if (!string.IsNullOrEmpty(CurrencyName))
            {
                string strSql = "SELECT n_ID FROM dbo.TCode_Currency WHERE s_Name='" + CurrencyName + "' OR s_CurrencyCode='" + CurrencyName + "'";
                return _dbHelper.GetbySql(strSql, commDB, _connection);
            }
            return -1;
        }

        private void AddAgencyRess(DataRow _dr, int nAgencyID, string type, string commDB, SqlConnection _connection)
        {
            _dealingClientData.GetIDbyClientAddress(type, nAgencyID, commDB, _connection);

            int nCountry = _dealingClientData.GetAddressIDByName(_dr["国家1"].ToString().Trim().Replace("'", "''"), commDB, _connection);
            int nCountry2 = _dealingClientData.GetAddressIDByName(_dr["国家2"].ToString().Trim().Replace("'", "''"), commDB, _connection);
            int nCountry3 = _dealingClientData.GetAddressIDByName(_dr["国家3"].ToString().Trim().Replace("'", "''"), commDB, _connection);

            string sAddress = _dealingClientData.IPtype(_dr["地址1"].ToString().Trim().Replace("'", "''"));
            string sAddress2 = _dealingClientData.IPtype(_dr["地址2"].ToString().Trim().Replace("'", "''"));
            string sAddress3 = _dealingClientData.IPtype(_dr["地址3"].ToString().Trim().Replace("'", "''"));

            AddAgencyAddress(nAgencyID, nCountry, _dr["省/州、自治区、直辖市1"].ToString().Trim().Replace("'", "''"), _dr["市县1"].ToString().Trim().Replace("'", "''"), _dr["街道门牌1"].ToString().Trim().Replace("'", "''"), _dr["邮编1"].ToString().Trim().Replace("'", "''"), sAddress, commDB, _connection);
            AddAgencyAddress(nAgencyID, nCountry2, _dr["省/州、自治区、直辖市2"].ToString().Trim().Replace("'", "''"), _dr["市县2"].ToString().Trim().Replace("'", "''"), _dr["街道门牌2"].ToString().Trim().Replace("'", "''"), _dr["邮编2"].ToString().Trim().Replace("'", "''"), sAddress2, commDB, _connection);
            AddAgencyAddress(nAgencyID, nCountry3, _dr["省/州、自治区、直辖市3"].ToString().Trim().Replace("'", "''"), _dr["市县3"].ToString().Trim().Replace("'", "''"), _dr["街道门牌3"].ToString().Trim().Replace("'", "''"), _dr["邮编3"].ToString().Trim().Replace("'", "''"), sAddress3, commDB, _connection);

        }

        private void AddAgencyAddress(int n_ClientID, int nCountry, string sState, string sCity, string s_Street, string s_ZipCode, string stype, string commDB, SqlConnection _connection)
        {
            string strSql = "INSERT INTO dbo.TCstmr_AgencyAddress( n_AgencyID ,n_Country ,s_State ,s_City , s_Street ,s_ZipCode,s_Type)" +
                           "VALUES(" + n_ClientID + "," + nCountry + ",'" + sState + "','" + sCity + "','" + s_Street + "','" + s_ZipCode + "','" + stype + "')";

            _dbHelper.InsertbySql(strSql, 0, commDB, _connection);
        }

        #endregion

        #region 外代理信息-联系人
        public int AddAgencyContact(DataRow _dr, int _row,string commDB, SqlConnection _connection)
        {

            int nAgencyID = _dealingClientData.GetClientIDByCaseSerialandCode(_dr["案件编号"].ToString(), "Agency", _dr["外代理"].ToString(), commDB, _connection);
            int nCaseID = _dbHelper.GetIDbyName(_dr["案件编号"].ToString().Trim(), 2, _connection);
            if (nAgencyID > 0)
            {
                int nLanguage = _dealingClientData.GetLanguageIDByName(_dr["语言"].ToString().Trim().Replace("'", "''"), commDB, _connection);
                string sPhone = _dr["座机"].ToString().Trim().Replace("'", "''");
                string sMobile = _dr["手机"].ToString().Trim().Replace("'", "''");
                string sFirstName = _dr["名"].ToString().Trim().Replace("'", "''");
                string sIPType = _dr["委托业务"].ToString().Trim().Replace("'", "''");
                string sEmail = _dr["email"].ToString().Trim().Replace("'", "''");
                int nContactID = _dealingClientData.GetClientIDByCaseSerial(sFirstName, sEmail, nAgencyID, "Agency", commDB, _connection);
                string strSql = "";
                if (nContactID > 0)
                {
                    strSql = "update TCstmr_AgencyContact set s_Phone='" + sPhone + "',s_Mobile='" + sMobile + "',s_IPType='" + sIPType + "',n_Language=" + nLanguage + " where n_ContactID=" + nContactID;
                }
                else
                {
                    strSql = " INSERT INTO dbo.TCstmr_AgencyContact( n_AgencyID ,s_FirstName , s_IPType ,n_Language ,s_Phone ,s_Mobile,s_Email)" +
                        " VALUES  ( " + nAgencyID + ",'" + sFirstName + "','" + sIPType + "'," + nLanguage + ",'" + sPhone + "','" + sMobile + "','" + sEmail + "')";

                }

                if (_dbHelper.InsertbySql(strSql, 0, commDB,_connection) > 0)
                {
                    nContactID = _dealingClientData.GetClientIDByCaseSerial(sFirstName, sEmail, nAgencyID, "Agency", commDB, _connection);
                    _dealingClientData.AddRess(_dr, nContactID, "Agency", commDB, _connection);
                    if (nCaseID > 0)
                    {
                        return _dealingClientData.InsertTCaseContact(nCaseID, "Agency", nContactID, "", _row, commDB, _connection); 
                    }
                }
            }
            else
            {
                _dbHelper.InsertLog(nCaseID, _dr["案件编号"].ToString(), _row, "外代理", "外代理-联系人-" + _row, "未找存有联系人案件:" + _dr["案件编号"].ToString() +":"+ _dr["名"].ToString() + _dr["email"].ToString(), "", commDB, _connection);
            }
            return 0;
        }
        #endregion

        #region 外代理信息-财务
        public int AddAgencyBill(DataRow _dr,int _row,string commDB,SqlConnection _connection)
        {
            int nAgencyID = _dealingClientData.GetClientandApplicantIDByName(_dr["代码"].ToString(), "CoopAgency", commDB, _connection);
            if (string.IsNullOrEmpty(_dr["代码"].ToString()))
            {
                nAgencyID = _dealingClientData.GetClientandApplicantIDByName(_dr["名称"].ToString(), "CoopAgencyName", commDB, _connection);
            }
            if (nAgencyID > 0)
            {
                string strSql = "update TCstmr_CoopAgency set s_BeneficiaryBankName='" + _dr["收款银行名称"].ToString() + "',s_BeneficiaryBankAddress='" + _dr["收款银行地址"].ToString() + "',s_BeneficiaryName='" + _dr["收款人名称"].ToString() + "'" +
                                ",s_BeneficiaryAddress='" + _dr["收款人地址"].ToString() + "',s_BeneficiaryAccountNumber='" + _dr["收款人账户号码"].ToString() + "',s_IBAN='" + _dr["IBAN"].ToString() + "',s_SwiftCode='" + _dr["Swift Code"].ToString() + "'" +
                                ",S_ABA='" + _dr["ABA"].ToString() + "'   where n_AgencyID =" + nAgencyID;
                return _dbHelper.InsertbySql(strSql, 0, commDB, _connection);
            }
            else
            {
                _dbHelper.InsertLog(0, _dr["代码"].ToString() + ":" + _dr["名称"].ToString(), _row, "外代理", "外代理-财务-" + _row, "未找外代理信息,编号：" + _dr["代码"].ToString() + ":" + _dr["名称"].ToString(), "", commDB, _connection); 
            }
            return 0;
        }
        #endregion
    }
}
