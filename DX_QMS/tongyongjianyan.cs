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
    public partial class tongyongjianyan : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public string oldtestitem = "";
        private IQC ic = new IQC(); 
        public TestBigTypeSet TBS;
        public TestItemInfo TII;
        public TestSubItemInfo TSI;
        // public TestProgSet TPS;
        public TestToolsSet TTS;
        public ReliabilitySet RS;
        public IQCTestCustomer TC;
        public tongyongjianyan()
        {
            InitializeComponent();
            setRule();
            setStyleForCurrentForm(this);

        }
        private void setStyleForCurrentForm(Form frm)
        {
            frm.BackColor = Color.FromArgb(210, 220, 230);
            foreach (Control con in frm.Controls)
            {
                if (con is TableLayoutPanel)
                {
                    foreach (Control contrl in con.Controls)
                        setControlStyle(contrl);
                }
                else if (con is GroupBox)
                {
                    foreach (Control contrl in con.Controls)
                        setControlStyle(contrl);
                }
                else if (con is TabControl)
                {
                    foreach (TabPage tp in ((TabControl)con).TabPages)
                    {
                        tp.BackColor = frm.BackColor;
                        foreach (Control contrl in tp.Controls)
                        {
                            setControlStyle(contrl);
                        }
                    }
                }
                else
                {
                    setControlStyle(con);
                }
            }
        }

        private void setControlStyle(Control contrl)
        {
            if (contrl is DataGridView)
            {
                ((DataGridView)contrl).ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                ((DataGridView)contrl).ColumnHeadersDefaultCellStyle.BackColor = Color.LightBlue;
                ((DataGridView)contrl).RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                ((DataGridView)contrl).RowsDefaultCellStyle.SelectionBackColor = Color.LightBlue;
                ((DataGridView)contrl).RowsDefaultCellStyle.SelectionForeColor = Color.Black;
                ((DataGridView)contrl).AllowUserToResizeRows = false;
            }
            else if (contrl is Button)
            {
                ((Button)contrl).Image = (Image)Properties.Resources.blue;
               
            }
        }

        private void btnTestType_Click(object sender, EventArgs e)
        {
            if (TBS == null || TBS.IsDisposed)
            {
                TBS = new TestBigTypeSet();
                TBS.ShowDialog();
            }
            else
            {
                TBS.Activate();
                TBS.ShowDialog();
            }
        }

        private void btntool_Click(object sender, EventArgs e)
        {
            if (TTS == null || TTS.IsDisposed)
            {
                TTS = new TestToolsSet();
                TTS.ShowDialog();
            }
            else
            {
                TTS.Activate();
                TTS.ShowDialog();
            }
        }

        private void btnTestItem_Click(object sender, EventArgs e)
        {
            if (TII == null || TII.IsDisposed)
            {
                TII = new TestItemInfo();
                TII.ShowDialog();
            }
            else
            {
                TII.Activate();
                TII.ShowDialog();
            }
        }

        private void btnTestSubItem_Click(object sender, EventArgs e)
        {
            if (TSI == null || TSI.IsDisposed)
            {
                TSI = new TestSubItemInfo();
                TSI.ShowDialog();
            }
            else
            {
                TSI.Activate();
                TSI.ShowDialog();
            }
        }

        private void btnprogset_Click(object sender, EventArgs e)
        {

        }

        private void btnreliability_Click(object sender, EventArgs e)
        {
            if (RS == null || RS.IsDisposed)
            {
                RS = new ReliabilitySet();
                RS.ShowDialog();
            }
            else
            {
                RS.Activate();
                RS.ShowDialog();
            }
        }

        private void BtnCutomer_Click(object sender, EventArgs e)
        {
            if (TC == null || TC.IsDisposed)
            {
                TC = new IQCTestCustomer();
                TC.ShowDialog();
            }
            else
            {
                TC.Activate();
                TC.ShowDialog();
            }
        }

        private void setRule()
        {
            string post = "";
            if (Login.manager != "")
            {
                post = Login.manager;
            }
            else
            {
                post = Login.post;
            }
           Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "通用设置");
            // btnSelect.Enabled = bool.Parse(GroupPermission.conversion(dic["hasQuery"].ToString()));
            btnSelect.Enabled = bool.Parse(dic["hasQuery"].ToString());
            btnAdd.Enabled = bool.Parse(GroupPermission.conversion(dic["hasInsert"].ToString()));
            btnDelete.Enabled = bool.Parse(GroupPermission.conversion(dic["hasDelete"].ToString()));
            btnupdate.Enabled = bool.Parse(GroupPermission.conversion(dic["hasUpdate"].ToString()));
        }
        private void tongyongjianyan_Load(object sender, EventArgs e)
        {
            bindDeviceType();
            txtTestItem.Select();
            setRule();
            //dgvDefect.AutoGenerateColumns = false;
            dgvDefect.DataSource = null;
            if (this.gridView.RowCount > 0)
            cbTestType.SelectedIndex = 0;
        }
        private void bindDeviceType()
        {
            string sql = "select checkType from OQC_CheckType";
            DataTable dt = Common.DbAccess.SelectBySql(sql).Tables[0];
            cbTestType.DataSource = dt;
            cbTestType.DisplayMember = dt.Columns["checkType"].ToString();
            cbTestType.ValueMember = dt.Columns["checkType"].ToString();
            this.btnSelect_Click(null, null);
        }

        private DataTable GetType(string checktype, string stype)
        {
            string sql = "select Definetype,Definevalue,sort,code from OQC_TypeDefine where Definetype='" + checktype + "' and Definevalue like '%" + stype + "%'";
            DataTable dt = Common.DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        private void btnSelect_Click(object sender, EventArgs e)
        {
            if (cbTestType.SelectedValue == null) return;
            DataTable dt = GetType(cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim());
            if (dt.Rows.Count > 0)
                dgvDefect.DataSource = dt;
            else
                dgvDefect.DataSource = null;
        }

        private void cbTestType_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnSelect_Click(sender, e);
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            dgvDefect.DataSource = null;
            if (cbTestType.Enabled)
            cbTestType.SelectedIndex = -1;
            txtTestItem.Text = "";
            txtid.Text = "";
            oldtestitem = "";
            txtconment.Text = "";
            this.txtTestItem.BackColor = Color.White;
            this.btnAdd.Text = "新 增";
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (txtTestItem.Text.Trim() == "") return;
            int m = 0;
            if (!int.TryParse(txtid.Text, out m))
            {
                MessageBox.Show("顺序请输入数字");
                return;
            }
            int i = ic.AddOQCCheckItem("新增", cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim(), int.Parse(txtid.Text), oldtestitem, txtconment.Text);
            if (i > 0)
                dgvDefect.DataSource = GetType(cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim());
            oldtestitem = "";
            txtTestItem.BackColor = Color.White;
            btnAdd.Text = "新 增";
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            //if (this.gridView.RowCount > 0)
            //{
            //    if (MessageBox.Show("确定删除选中的测试项目？", "删除提示！", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            //    {
            //        int temp = ic.AddOQCCheckItem("删除", gridView.GetRowCellValue(gridView.FocusedRowHandle, "类别").ToString(), gridView.GetRowCellValue(gridView.FocusedRowHandle, "类别名称").ToString(), 0, "", "");
            //        if (temp > 0)
            //            dgvDefect.DataSource = GetType(cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim());
            //    }
            //}

            DataTable de = dgvDefect.DataSource as DataTable;
            if (de == null || de.Rows.Count < 0)
                return;
            if (gridView.FocusedRowHandle < 0)
                return;       
                if (MessageBox.Show("确定删除选中的测试类别？", "删除提示！", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    int temp = ic.AddOQCCheckItem("删除", gridView.GetFocusedRowCellValue("Definetype").ToString(),gridView.GetFocusedRowCellValue("Definevalue").ToString(), 0, "", "");
                    if (temp > 0)
                        dgvDefect.DataSource = GetType(cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim());
                }          
        }

        private void btnupdate_Click(object sender, EventArgs e)
        {
            if (txtTestItem.Text.Trim() == "") return;
            int m = 0;
            if (!int.TryParse(txtid.Text, out m))
            {
                MessageBox.Show("顺序请输入数字");
                return;
            }
            int i = ic.AddOQCCheckItem("更新", cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim(), int.Parse(txtid.Text), oldtestitem, txtconment.Text);
            if (i > 0)
                dgvDefect.DataSource = GetType(cbTestType.SelectedValue.ToString(), txtTestItem.Text.Trim());
            oldtestitem = "";
            txtTestItem.BackColor = Color.White;
            this.btnupdate.Enabled = false;
        }

        private void dgvDefect_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //gridView1.GetFocusedRowCellValue("code").ToString();
            cbTestType.SelectedValue = gridView.GetFocusedRowCellValue("Definetype").ToString();
            txtTestItem.Text = gridView.GetFocusedRowCellValue("Definevalue").ToString();
            txtid.Text = gridView.GetFocusedRowCellValue("sort").ToString();
            oldtestitem = gridView.GetFocusedRowCellValue("Definevalue").ToString();
            this.txtconment.Text = gridView.GetFocusedRowCellValue("code").ToString();
            this.txtTestItem.BackColor = Color.Yellow; 
            this.btnupdate.Enabled = true;
        }

        private void gridView_Click(object sender, EventArgs e)
        {
            /////gridView.GetFocusedRowCellValue("产品编码").ToString();
            DataTable de = dgvDefect.DataSource as DataTable;
            if (de == null || de.Rows.Count < 0)
                return;
            cbTestType.Text = gridView.GetFocusedRowCellValue("Definetype").ToString();
            txtTestItem.Text = gridView.GetFocusedRowCellValue("Definevalue").ToString();
            txtid.Text = gridView.GetFocusedRowCellValue("sort").ToString();
            txtconment.Text = gridView.GetFocusedRowCellValue("code").ToString();
            oldtestitem = gridView.GetFocusedRowCellValue("Definevalue").ToString();
            this.txtTestItem.BackColor = Color.Yellow;
            this.btnupdate.Enabled = true;
        }

    }
}