using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DX_QMS.Common;

namespace DX_QMS
{
    public partial class ReliabilityTypeSet : DevExpress.XtraEditors.XtraForm
    {
        private IQC ic = new IQC();
        public ReliabilityTypeSet()
        {
            InitializeComponent();
        }

        private void ReliabilityTypeSet_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
            this.Close();
        }
        private void bindTypeSet(string testtype, string tools)
        {
            DataSet ds = ic.SelectTestTypeRecord("查询", testtype, tools, "");
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                databind.DataSource = ds.Tables[0];
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (txtTestType.Text.Trim() == "") return;
            if (txttestbigtype.Text.Trim() == "") return;

            int upTemp = ic.AddNewTestTypeRecord("新增", txtTestType.Text.Trim(), txttestbigtype.Text, txttestcontent.Text);
            if (upTemp > 0)
                MessageBox.Show("新增成功！", "修改提示！", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(txtTestType.Text + "存在！", "修改提示！", MessageBoxButtons.OK, MessageBoxIcon.Error);
            txtTestType.Text = "";
            txtTestType.Focus();
            bindTypeSet(txtTestType.Text.Trim(), txttestbigtype.Text);
        }

        private void btndel_Click(object sender, EventArgs e)
        {
            ///// string customer = gridView.GetFocusedRowCellValue("customer").ToString();

            if (gridView.FocusedRowHandle < 0)
                return;

            if (MessageBox.Show("确定要删除选中项吗？", "提醒", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                
                int m = 0;
                for (int i = gridView.GetSelectedRows().Length; i > 0; i--)    //// DataRow dr = gridView1.GetDataRow(i);
                {
                    DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[i-1]);
                    ////   dr["TestType"].ToString ();TTypes
                    int upTemp = ic.AddNewTestTypeRecord("删除", dr["TestType"].ToString(), dr["TTypes"].ToString(), "");
                    if (upTemp > 0)
                    {
                        m++;
                        gridView.DeleteRow(gridView.GetSelectedRows()[i - 1]);
                        ////this.databind.Rows.RemoveAt(databind.SelectedRows[i - 1].Index);
                    }

                }
                if (m > 0)
                    MessageBox.Show(m.ToString() + "项,删除成功！", "修改提示！", MessageBoxButtons.OK, MessageBoxIcon.Information);
               

                txtTestType.Text = "";
                txtTestType.Focus();
                bindTypeSet(txtTestType.Text.Trim(), txttestbigtype.Text);
            }
        }

        private void btnquery_Click(object sender, EventArgs e)
        {
            bindTypeSet(txtTestType.Text.Trim(), this.txttestbigtype.Text);
        }
    }
}