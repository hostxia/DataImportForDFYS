using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Import
{
    public partial class Frm_AddData : Form
    {
        /// <summary>
        /// excel中的数据集合
        /// </summary>
        private DataTable dtExcel;
        public Frm_AddData()
        {
            InitializeComponent();
        }

        private void bt_ok_Click(object sender, EventArgs e)
        {

        }
       
        private void ToolImportCaseData()
        { 
            #region 基本信息
            #endregion

            #region 地址
            #endregion


            #region 联系人
            #endregion

            #region 财务信息
            #endregion
        }
    }
}
