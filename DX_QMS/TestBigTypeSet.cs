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
    public partial class TestBigTypeSet : DevExpress.XtraEditors.XtraForm
    {
        private IQC ic = new IQC();
        public TestBigTypeSet()
        {
            InitializeComponent();
        }

        private void TestBigTypeSet_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
            this.Close();
        }


        private void bindTypeSet(string testtype, string tools)
        {
            DataSet ds = ic.SelectTestTypeRecord("查询", testtype, "测试类别", "");
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                databind_father.DataSource = ds.Tables[0];
            }
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (txtTestType.Text.Trim() == "") return;

            int upTemp = ic.AddNewTestTypeRecord("新增", txtTestType.Text.Trim(), "测试类别", "");
            if (upTemp > 0)
                MessageBox.Show("新增成功！", "修改提示！", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(txtTestType.Text + "存在！", "修改提示！", MessageBoxButtons.OK, MessageBoxIcon.Error);
            txtTestType.Text = "";
            txtTestType.Focus();
            bindTypeSet(txtTestType.Text.Trim(), "测试类别");
        }

        private void btndel_Click(object sender, EventArgs e)
        {
            if (txtTestType.Text.Trim() == "") return;

            int upTemp = ic.AddNewTestTypeRecord("删除", txtTestType.Text.Trim(), "测试类别", "");
            if (upTemp > 0)
                MessageBox.Show("删除成功！", "修改提示！", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(txtTestType.Text + "不存在！", "修改提示！", MessageBoxButtons.OK, MessageBoxIcon.Error);
            txtTestType.Text = "";
            txtTestType.Focus();

            bindTypeSet(txtTestType.Text.Trim(), "测试类别");
        }

        private void btnquery_Click(object sender, EventArgs e)
        {
            bindTypeSet(txtTestType.Text.Trim(), "测试类别");
        }
        private void btnupdate_Click(object sender, EventArgs e)
        {
            int upTemp = ic.AddNewTestTypeRecord("更新", txtTestType.Text.Trim(), "测试类别", oldtesttype);
            if (upTemp > 0)
            {
                oldtesttype = "";
                MessageBox.Show("更新成功！", "修改提示！", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnupdate.Enabled = false;
                btnquery_Click(sender, e);
            }
            else
                MessageBox.Show(txtTestType.Text + "更新失败！", "修改提示！", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        public string oldtesttype = "";

        private void btnquery_Click_1(object sender, EventArgs e)
        {
            bindTypeSet(txtTestType.Text.Trim(), "测试类别");
        }

        private void databind_father_Click(object sender, EventArgs e)
        {
            txtTestType.Text = databind.GetFocusedRowCellValue("TestType").ToString();
            oldtesttype = databind.GetFocusedRowCellValue("TestType").ToString();
            this.btnupdate.Enabled = true;
        }
    }
}