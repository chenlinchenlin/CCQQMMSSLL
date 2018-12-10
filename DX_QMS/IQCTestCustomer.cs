using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DX_QMS.Common;
using DevExpress.XtraBars;

namespace DX_QMS
{
    public partial class IQCTestCustomer : DevExpress.XtraEditors.XtraForm
    {
        public IQCTestCustomer()
        {
            InitializeComponent();
        }

        private void IQCTestCustomer_Load(object sender, EventArgs e)
        {
            txtcustometype.SelectedIndex = 0;
        }

        private void txtcustometype_SelectedIndexChanged(object sender, EventArgs e)
        {
            sBtnselect_Click(sender,e);
        }

        private void sBtnselect_Click(object sender, EventArgs e)
        {
            string where = " where 1=1 ";
            string custometype = txtcustometype.Text;
            string customer = txtcustomer.Text;

            if (!string.IsNullOrEmpty(custometype))
            {
                where += " and custometype = '" + custometype + "' ";
            }
            if (!string.IsNullOrEmpty(customer))
            {
                where += " and customer = '" + customer + "' ";
            }

            string sql = @"  select custometype 客户类别,customer 客户,remark 备注,updateman 更新人,updatetime 更新时间 from IQC_Customer  ";
            sql += where + " order by updatetime desc ";

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

            if (dt != null && dt.Rows.Count > 0)
            {
                gridControl.DataSource = dt;
            }
            else
            {
                gridControl.DataSource = null;
            }

        }

        private void sBtnadd_Click(object sender, EventArgs e)
        {
            DataTable dt = null;
            string sql = @"  select 1 from IQC_Customer where custometype = '"+ txtcustometype.Text+ "' and  customer = '"+ txtcustomer.Text+ "'  ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                MessageBox.Show("该客户已经存在","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information );
                return;
            }
            sql = "  insert into IQC_Customer ( custometype,customer,remark,updateman,updatetime)	values ( '"+txtcustometype.Text+ "','"+txtcustomer.Text+"','"+ txtremark.Text+ "','"+Login.username+"',GETDATE()) ";
            bool flat =  DbAccess.ExecuteSql(sql);
            if (flat == true)
            {
                MessageBox.Show("新增成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                sBtnselect_Click(sender,e);
            }
            else
            {
                MessageBox.Show("新增失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void sBtndelete_Click(object sender, EventArgs e)
        {
            if (gridView.RowCount < 1)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;

            string custometype = gridView.GetFocusedRowCellValue("客户类别").ToString();
            string customer = gridView.GetFocusedRowCellValue("客户").ToString();

            string sql = @" delete IQC_Customer where custometype = '"+custometype+ "' and  customer = '"+customer+ "'   ";
            bool flat = DbAccess.ExecuteSql(sql);
            if (flat == true)
            {
                MessageBox.Show("删除成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                sBtnselect_Click(sender, e);
            }
            else
            {
                MessageBox.Show("删除失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void sBtnreset_Click(object sender, EventArgs e)
        {
            txtcustometype.Text = "";
            txtcustomer.Text = "";
            txtremark.Text = "";
            gridControl.DataSource = null;
        }

        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {

            txtcustometype.Text = gridView.GetFocusedRowCellValue("客户类别").ToString();
            txtcustomer.Text = gridView.GetFocusedRowCellValue("客户").ToString();
            txtremark.Text = gridView.GetFocusedRowCellValue("备注").ToString();

        }
    }
}