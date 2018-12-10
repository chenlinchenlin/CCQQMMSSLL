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
    public partial class TestToolsSet : DevExpress.XtraEditors.XtraForm
    {
        private IQC ic = new IQC();
        public TestToolsSet()
        {
            InitializeComponent();
        }

        private void TestToolsSet_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
            this.Close();
        }

        private void TestToolsSet_Load(object sender, EventArgs e)
        {
            ifyiqi.SelectedIndex = 0;
        }

        private void bindTypeSet(string testtype, string tools, string yiqi)
        {
            DataSet ds = ic.SelectTestTypeRecord("查询", testtype, "测试工具", yiqi);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                databind.DataSource = ds.Tables[0];
            }
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (txtTestType.Text.Trim() == "") return;

            int upTemp = ic.AddNewTestTypeRecord("新增", txtTestType.Text.Trim(), "测试工具", ifyiqi.Text);
            if (upTemp > 0)
                MessageBox.Show("新增成功！", "修改提示！", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(txtTestType.Text + "存在！", "修改提示！", MessageBoxButtons.OK, MessageBoxIcon.Error);
            txtTestType.Text = "";
            txtTestType.Focus();

            bindTypeSet(txtTestType.Text.Trim(), "测试工具", ifyiqi.Text);
        }

        private void btndel_Click(object sender, EventArgs e)
        {
            if (txtTestType.Text.Trim() == "") return;

            int upTemp = ic.AddNewTestTypeRecord("删除", txtTestType.Text.Trim(), "测试工具", "");
            if (upTemp > 0)
                MessageBox.Show("删除成功！", "修改提示！", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(txtTestType.Text + "不存在！", "修改提示！", MessageBoxButtons.OK, MessageBoxIcon.Error);
            txtTestType.Text = "";
            txtTestType.Focus();

            bindTypeSet(txtTestType.Text.Trim(), "测试工具", ifyiqi.Text);
        }

        private void btnquery_Click(object sender, EventArgs e)
        {
            bindTypeSet(txtTestType.Text.Trim(), "测试工具", ifyiqi.Text);
        }
    }
}