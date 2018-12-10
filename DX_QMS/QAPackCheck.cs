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
    public partial class QAPackCheck : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private string lotno = "";
        public QAPackCheck()
        {
            InitializeComponent();
        }

        private void txtsn_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string sql = "select lotno from Pack_Maintain where serialno='" + txtsn.Text.Trim() + "'";
                DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                if (dt.Rows.Count > 0)
                {
                    txtcolorsn.Text = "";
                    txtcolorsn.Focus();
                    lotno = dt.Rows[0]["lotno"].ToString();
                    txtsn.Enabled = false;
                }
                else
                {
                    lblinfo.Text = "机身号:" + txtsn.Text + ",不存在!";
                    lblinfo.ForeColor = Color.Red;
                    txtsn.Text = "";
                    txtsn.Focus();
                    txtsn.Enabled = true;
                    lotno = "";
                }
            }
        }

        private void txtcolorsn_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txtcolorsn.Text != txtsn.Text)
                {
                    lblinfo.Text = txtcolorsn.Text + "彩合号与机身号不一致!";
                    lblinfo.ForeColor = Color.Red;
                    txtcolorsn.Text = "";
                    txtcolorsn.Focus();
                    txtcolorsn.Enabled = true;
                }
                else
                {
                    txtboxsn.Text = "";
                    txtboxsn.Focus();
                    txtcolorsn.Enabled = false;
                }
            }
        }


        void Search( string SN )
        {
            string sql = @" select SN,CorSN,LotNo from IQC_QACheck where SN='"+ SN + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                databind.DataSource = dt;
            }
            else
           {
                MessageBox.Show("没有数据","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);

            }
        }

        public DataTable CreateTable()
        {
            DataTable namesTable = new DataTable("SNbaddetails");
            DataColumn SN = new DataColumn();
            SN.DataType = System.Type.GetType("System.String");
            SN.ColumnName = "SN";
            namesTable.Columns.Add(SN);

            DataColumn CorSN = new DataColumn();
            CorSN.DataType = System.Type.GetType("System.String");
            CorSN.ColumnName = "CorSN";
            namesTable.Columns.Add(CorSN);

            DataColumn LotNo = new DataColumn();
            LotNo.DataType = System.Type.GetType("System.String");
            LotNo.ColumnName = "LotNo";
            namesTable.Columns.Add(LotNo);

            return namesTable;
        }


        private void txtboxsn_KeyUp(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Enter)
            {
                if (txtboxsn.Text.Trim().ToUpper() != lotno.ToUpper())
                {
                    lblinfo.Text = txtboxsn.Text + "外箱号不正确,正确为:" + lotno;
                    lblinfo.ForeColor = Color.Red;
                    txtboxsn.Enabled = true;
                    txtboxsn.Text = "";
                    txtboxsn.Focus();
                }
                else
                {
                    string sql = "if not exists(select 1 from IQC_QACheck where SN='" + txtsn.Text + "'" + ")";
                    sql += " insert into IQC_QACheck(SN,CorSN,LotNo,operuser,operdate) values('" + txtsn.Text + "','" + txtcolorsn.Text + "','" + txtboxsn.Text + "','" + Login.userId + "',getdate()" + ")";
                    bool F = DbAccess.ExecuteSql(sql);
                    if (F)
                    {
                        //databind.Rows.Add(1);
                        //int count = databind.Rows.Count;
                        //databind.Rows[count - 1].Cells["SN"].Value = txtsn.Text;
                        //databind.Rows[count - 1].Cells["ColorSN"].Value = txtcolorsn.Text;
                        //databind.Rows[count - 1].Cells["BoxSN"].Value = txtboxsn.Text;
                        Search(txtsn.Text);
                        lblinfo.Text = txtsn.Text + "核对成功";
                        lblinfo.ForeColor = Color.Blue;
                        txtboxsn.Text = "";
                        txtcolorsn.Text = "";
                        txtcolorsn.Enabled = true;
                        txtsn.Text = "";
                        txtsn.Enabled = true;
                        txtsn.Focus();
                    }
                    else
                    {
                        DataTable de = databind.DataSource as DataTable;


                        for (int k = 0; k < de.Rows.Count; k++)
                        {
                            if (de.Rows[k]["SN"].ToString() == txtsn.Text.ToUpper().Trim())
                            {
                                lblinfo.Text = txtsn.Text + "已经核对过";
                                lblinfo.ForeColor = Color.Blue;
                                txtboxsn.Text = "";
                                txtcolorsn.Text = "";
                                txtcolorsn.Enabled = true;
                                txtsn.Text = "";
                                txtsn.Enabled = true;
                                txtsn.Focus();
                                return;
                            }
                        }

                        Search(txtsn.Text);
                        lblinfo.Text = txtsn.Text + "已经核对过";
                        lblinfo.ForeColor = Color.Blue;
                        txtboxsn.Text = "";
                        txtcolorsn.Text = "";
                        txtcolorsn.Enabled = true;
                        txtsn.Text = "";
                        txtsn.Enabled = true;
                        txtsn.Focus();

                        /*
                        databind.Rows.Add(1);
                        int count = databind.Rows.Count;
                        databind.Rows[count - 1].Cells["SN"].Value = txtsn.Text;
                        databind.Rows[count - 1].Cells["ColorSN"].Value = txtcolorsn.Text;
                        databind.Rows[count - 1].Cells["BoxSN"].Value = txtboxsn.Text;
                        lblinfo.Text = txtsn.Text + "已经核对过";
                        lblinfo.ForeColor = Color.Blue;
                        txtboxsn.Text = "";
                        txtcolorsn.Text = "";
                        txtcolorsn.Enabled = true;
                        txtsn.Text = "";
                        txtsn.Enabled = true;
                        txtsn.Focus();
                        */


                    }

                }
            }
        }
        private void btnsearch_Click(object sender, EventArgs e)
        {
            if (txtsn.Text == "" && txtcolorsn.Text == "" && txtboxsn.Text == "" && txtworkno.Text == "") return;
            //databind.Rows.Clear();
            databind.DataSource = null;
            string sql = "select SN,CorSN,q.LotNo,workno from IQC_QACheck q inner join Pack_Maintain p on q.SN=p.serialno where 1=1";
            if (txtsn.Text != "")
                sql += " and SN='" + this.txtsn.Text + "'";
            if (txtcolorsn.Text != "")
                sql += " and CorSN='" + txtcolorsn.Text + "'";
            if (txtboxsn.Text != "")
                sql += " and q.LotNo='" + txtboxsn.Text + "'";
            if (this.txtworkno.Text != "")
                sql += " and workno='" + txtworkno.Text + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                databind.DataSource = dt;
            }
            else
            {
                MessageBox.Show("没有数据","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    databind.Rows.Add(1);
            //    int count = databind.Rows.Count;
            //    databind.Rows[count - 1].Cells["SN"].Value = dt.Rows[i]["SN"].ToString();
            //    databind.Rows[count - 1].Cells["ColorSN"].Value = dt.Rows[i]["CorSN"].ToString();
            //    databind.Rows[count - 1].Cells["BoxSN"].Value = dt.Rows[i]["LotNo"].ToString();
            //}
        }
    }
}