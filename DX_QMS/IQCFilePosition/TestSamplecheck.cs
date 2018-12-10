using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DX_QMS.Common;

namespace DX_QMS.IQCFilePosition
{
    public partial class TestSamplecheck : DevExpress.XtraEditors.XtraForm
    {

        public string ifcorrect = "";
        public TestSamplecheck()
        {
            InitializeComponent();
        }

        public TestSamplecheck(string productcode)
        {
            InitializeComponent();
            txtproductcode.Text = productcode;
        }


        private void TestSamplecheck_Load(object sender, EventArgs e)
        {
            string sql = " select supplier 供应商,position 样品位置 from IQC_TestSamplePosition  where productcode = '"+ txtproductcode.Text+ "' order by eventtime desc  ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                string simpleposition = "";
                for (int i = 0; i< dt.Rows.Count; i++ )
                {
                    simpleposition += " 【 供应商：" + dt.Rows[i]["供应商"].ToString() + "；样品位置：" + dt.Rows[i]["样品位置"].ToString()+ "】 ";
                }
                lblsimpleposition.Text = simpleposition;
                lblsimpleposition.ForeColor = Color.Red;
            }

        }

        private void txtsampleCode_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtsampleCode.Text != "")
            {
                txtsampleCode_Leave(sender, e);

                //string sql = "  select sampleCode from IQC_TestSamplePosition   where productcode = '"+txtproductcode.Text+ "' order by eventtime desc    ";
                //DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
                //if (dt != null && dt.Rows.Count > 0)
                //{
                //    string falt = "";
                //    for (int i = 0; i < dt.Rows.Count; i++)
                //    {
                //        string samplecode = dt.Rows[i]["sampleCode"].ToString();
                //        if (samplecode == txtsampleCode.Text)
                //        {
                //            lblinfo.Text = "虚拟编码正确";
                //            falt = lblinfo.Text;
                //            MessageBox.Show(lblinfo.Text, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //            break;
                //        }
                //    }
                //    if (falt != "虚拟编码正确")
                //    {
                //        MessageBox.Show("虚拟编码错误", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //    }
                //}
                //else
                //{
                //    MessageBox.Show("虚拟编码不存在", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //}
            }
        }

        private void txtsampleCode_Leave(object sender, EventArgs e)
        {
            string sql = "  select sampleCode from IQC_TestSamplePosition  where productcode = '" + txtproductcode.Text + "' order by eventtime desc    ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                string falt = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string samplecode = dt.Rows[i]["sampleCode"].ToString();
                    if (samplecode == txtsampleCode.Text)
                    {
                        lblinfo.Text = "虚拟编码正确";
                        falt = lblinfo.Text;
                        MessageBox.Show(lblinfo.Text, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    }
                }
                if (falt != "虚拟编码正确")
                {
                    MessageBox.Show("虚拟编码错误", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    ifcorrect = lblinfo.Text;
                    this.Dispose();
                    this.Close();
                }
            }
            else
            {
                MessageBox.Show("虚拟编码不存在", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        private void sBtnOK_Click(object sender, EventArgs e)
        {

            if (lblinfo.Text == "虚拟编码正确")
            {
                ifcorrect = lblinfo.Text;
                this.Dispose();
                this.Close();
            }

            //string sql = "  select sampleCode from IQC_TestSamplePosition   where productcode = '" + txtproductcode.Text + "' order by eventtime desc    ";
            //DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            //if (dt != null && dt.Rows.Count > 0)
            //{
            //    string falt = "";
            //    for (int i = 0; i < dt.Rows.Count; i++)
            //    {
            //        string samplecode = dt.Rows[i]["sampleCode"].ToString();
            //        if (samplecode == txtsampleCode.Text)
            //        {
            //            lblinfo.Text = "虚拟编码正确";
            //            falt = lblinfo.Text;
            //            MessageBox.Show(lblinfo.Text, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //            break;
            //        }
            //    }
            //    if (falt != "虚拟编码正确")
            //    {
            //        MessageBox.Show("虚拟编码错误", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("虚拟编码不存在", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
        }
    }
}