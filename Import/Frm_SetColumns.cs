using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Winform_SqlBulkCopy
{
    public partial class Frm_SetColumns : Form
    {
        #region 全局变量
        /// <summary>
        /// 委托
        /// </summary>
        public Action<List<SqlBulkCopyColumnMapping>> DeleteSet;
        List<string> ExcelColumns, SqlColumns;
        #endregion

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="_ExcelColumns">Excel列</param>
        /// <param name="_SqlColumns">数据库列</param>
        public Frm_SetColumns(List<string> _ExcelColumns, List<string> _SqlColumns)
        {
            ExcelColumns = _ExcelColumns;
            SqlColumns = _SqlColumns;
            InitializeComponent();
            Load += new EventHandler(SetColumns_Load);
        }
        #endregion

        #region 窗体加载
        /// <summary>
        /// 窗体加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void SetColumns_Load(object sender, EventArgs e)
        {
            Init();
            EventHand();
        }
        #endregion

        #region 初始化
        /// <summary>
        /// 初始化
        /// </summary>
        void Init()
        {
            MaximizeBox = false;
            MaximumSize = MinimumSize = Size;
            foreach (string item in ExcelColumns) listbox_Excel.Items.Add(item);
            foreach (string item in SqlColumns) listbox_Ssms.Items.Add(item);
            if (ExcelColumns.Count != SqlColumns.Count) Text += "(字段数量不匹配)";
        }
        #endregion

        #region 控件事件挂接
        /// <summary>
        /// 控件事件挂接
        /// </summary>
        void EventHand()
        {
            //Excel控制
            lb_Excel_First.Click += new EventHandler(lb_Click);
            lb_Excel_Before.Click += new EventHandler(lb_Click);
            lb_Excel_Next.Click += new EventHandler(lb_Click);
            lb_Excel_Last.Click += new EventHandler(lb_Click);
            lb_Excel_Del.Click += new EventHandler(lb_Click);
            //Sql控制
            lb_Sql_First.Click += new EventHandler(lb_Click);
            lb_Sql_Before.Click += new EventHandler(lb_Click);
            lb_Sql_Next.Click += new EventHandler(lb_Click);
            lb_Sql_Last.Click += new EventHandler(lb_Click);
            lb_DB_Del.Click += new EventHandler(lb_Click);
            //确定
            bt_OK.Click += new EventHandler(bt_OK_Click);
        }
        #endregion

        #region 控件事件响应
        void lb_Click(object sender, EventArgs e)
        {
            object excelitem = listbox_Excel.SelectedItem;
            object sqlitem = listbox_Ssms.SelectedItem;
            int excelindex, sqlindex;
            switch ((sender as Label).Name)
            {
                //Excel
                case "lb_Excel_First":
                    if (excelitem == null)
                    {
                        MessageBox.Show("请选择表格中字段");
                        return;
                    }
                    excelindex = listbox_Excel.SelectedIndex;
                    if (excelindex <= 0) return;
                    listbox_Excel.Items.Remove(listbox_Excel.SelectedItem);
                    listbox_Excel.Items.Insert(0, excelitem);
                    listbox_Excel.SelectedIndex = 0;
                    break;
                case "lb_Excel_Before":
                    if (excelitem == null)
                    {
                        MessageBox.Show("请选择表格中字段");
                        return;
                    }
                    excelindex = listbox_Excel.SelectedIndex;
                    if (excelindex <= 0) return;
                    listbox_Excel.Items.Remove(listbox_Excel.SelectedItem);
                    listbox_Excel.Items.Insert(excelindex - 1, excelitem);
                    listbox_Excel.SelectedIndex = excelindex - 1;
                    break;
                case "lb_Excel_Next":
                    if (excelitem == null)
                    {
                        MessageBox.Show("请选择表格中字段");
                        return;
                    }
                    excelindex = listbox_Excel.SelectedIndex;
                    if (excelindex >= listbox_Excel.Items.Count - 1) return;
                    listbox_Excel.Items.Remove(listbox_Excel.SelectedItem);
                    listbox_Excel.Items.Insert(excelindex + 1, excelitem);
                    listbox_Excel.SelectedIndex = excelindex + 1;
                    break;
                case "lb_Excel_Last":
                    if (excelitem == null)
                    {
                        MessageBox.Show("请选择表格中字段");
                        return;
                    }
                    excelindex = listbox_Excel.SelectedIndex;
                    if (excelindex >= listbox_Excel.Items.Count - 1) return;
                    listbox_Excel.Items.Remove(listbox_Excel.SelectedItem);
                    listbox_Excel.Items.Insert(listbox_Excel.Items.Count - 1, excelitem);
                    listbox_Excel.SelectedIndex = listbox_Excel.Items.Count - 1;
                    break;
                case "lb_Excel_Del":
                    if (excelitem == null)
                    {
                        MessageBox.Show("请选择表格中字段");
                        return;
                    }
                    excelindex = listbox_Excel.SelectedIndex;
                    if (excelindex >= listbox_Excel.Items.Count) return;
                    listbox_Excel.Items.Remove(listbox_Excel.SelectedItem);
                    listbox_Excel.SelectedIndex = listbox_Excel.Items.Count - 1;
                    break;
                //Sql
                case "lb_Sql_First":
                    if (sqlitem == null)
                    {
                        MessageBox.Show("请选择数据库中字段");
                        return;
                    }
                    sqlindex = listbox_Ssms.SelectedIndex;
                    if (sqlindex <= 0)
                    {
                        MessageBox.Show("请选择数据库中字段");
                        return;
                    }
                    listbox_Ssms.Items.Remove(listbox_Ssms.SelectedItem);
                    listbox_Ssms.Items.Insert(0, sqlitem);
                    listbox_Ssms.SelectedIndex = 0;
                    break;
                case "lb_Sql_Before":
                    if (sqlitem == null)
                    {
                        MessageBox.Show("请选择数据库中字段");
                        return;
                    }
                    sqlindex = listbox_Ssms.SelectedIndex;
                    if (sqlindex <= 0) return;
                    listbox_Ssms.Items.Remove(listbox_Ssms.SelectedItem);
                    listbox_Ssms.Items.Insert(sqlindex - 1, sqlitem);
                    listbox_Ssms.SelectedIndex = sqlindex - 1;
                    break;
                case "lb_Sql_Next":
                    if (sqlitem == null)
                    {
                        MessageBox.Show("请选择数据库中字段");
                        return;
                    }
                    sqlindex = listbox_Ssms.SelectedIndex;
                    if (sqlindex >= listbox_Ssms.Items.Count - 1) return;
                    listbox_Ssms.Items.Remove(listbox_Ssms.SelectedItem);
                    listbox_Ssms.Items.Insert(sqlindex + 1, sqlitem);
                    listbox_Ssms.SelectedIndex = sqlindex + 1;
                    break;
                case "lb_Sql_Last":
                    if (sqlitem == null)
                    {
                        MessageBox.Show("请选择数据库中字段");
                        return;
                    }
                    sqlindex = listbox_Ssms.SelectedIndex;
                    if (sqlindex >= listbox_Ssms.Items.Count - 1) return;
                    listbox_Ssms.Items.Remove(listbox_Ssms.SelectedItem);
                    listbox_Ssms.Items.Insert(listbox_Ssms.Items.Count, sqlitem);
                    listbox_Ssms.SelectedIndex = listbox_Ssms.Items.Count - 1; 
                    break; 
                case "lb_DB_Del":
                    if (sqlitem == null)
                    {
                        MessageBox.Show("请选择数据库中字段");
                        return;
                    }
                    sqlindex = listbox_Ssms.SelectedIndex;
                    if (sqlindex >= listbox_Ssms.Items.Count) return;
                    listbox_Ssms.Items.Remove(listbox_Ssms.SelectedItem);
                    listbox_Ssms.SelectedIndex = listbox_Ssms.Items.Count - 1;
                    break;
                default: break;
            }
        }

        void bt_OK_Click(object sender, EventArgs e)
        {
            if (listbox_Excel.Items.Count != listbox_Ssms.Items.Count)
            { MessageBox.Show("Excel表格与数据库的列不相同，不能保存!"); return; }
            List<SqlBulkCopyColumnMapping> list = new List<SqlBulkCopyColumnMapping>();
            for (int i = 0; i < listbox_Excel.Items.Count; i++)
                list.Add(new SqlBulkCopyColumnMapping(listbox_Excel.Items[i].ToString(), listbox_Ssms.Items[i].ToString()));
            DeleteSet(list);
            DialogResult = DialogResult.OK;
        }
        #endregion
    }
}
