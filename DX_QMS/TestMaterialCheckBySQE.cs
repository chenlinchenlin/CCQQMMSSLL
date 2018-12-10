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
    public partial class TestMaterialCheckBySQE : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public TestMaterialCheckBySQE()
        {
            InitializeComponent();
            setRule();
           // this.btnconfirm.Enabled = false;
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
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "标准");
            this.btnconfirm.Enabled = bool.Parse(dic["hasUpdate"].ToString());
        }

        private DataTable GetStateByProduct(string pcode)
        {
            string sql = "select States from  IQC_TestMaterialState where Materialcode='" + pcode + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        private void txtproductcode_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                this.lblinfo.Text = "";
                this.btnconfirm.Enabled = false;
                DataTable dt = GetStateByProduct(txtproductcode.Text);
                if (dt.Rows.Count > 0)
                {
                    txtcurrentcheck.Text = dt.Rows[0]["States"].ToString();
                    if (txtcurrentcheck.Text == "暂停检验")
                    {
                        txtmodifycheck.Text = "加严检验";
                        this.btnconfirm.Enabled = true;
                    }
                    else if (txtcurrentcheck.Text == "放宽检验")
                    {
                        txtmodifycheck.Text = "正常检验";
                        this.btnconfirm.Enabled = true;
                    }
                    else
                    {
                        this.lblinfo.Text = this.txtproductcode.Text + "当前检验标准为:" + txtcurrentcheck.Text + ",不能更改检验标准";
                        this.lblinfo.ForeColor = Color.Red;
                        txtcurrentcheck.Text = "";
                        txtproductcode.Text = "";
                        txtproductcode.Focus();
                    }
                }
                else
                {
                    this.lblinfo.Text = txtproductcode.Text + "编码不存在!";
                    this.lblinfo.ForeColor = Color.Red;
                    txtcurrentcheck.Text = "";
                    txtmodifycheck.Text = "";
                    txtproductcode.Text = "";
                    txtproductcode.Focus();
                }
            }

        }

        private void btnconfirm_Click(object sender, EventArgs e)
        {
            if (txtcurrentcheck.Text == "") return;
            DialogResult rt = MessageBox.Show("是否确定要更改检验标准?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (DialogResult.Yes == rt)
            {
                string sql = "update IQC_TestMaterialState set ReceptidOK='',ReceptidNG='',OKCount=0,NGCount=0,States='" + txtmodifycheck.Text + "',OperUser='" + Login.userId + "',OperDate=getdate() where Materialcode='" + txtproductcode.Text + "'";
                bool b = DbAccess.ExecuteSql(sql);
                if (b)
                {
                    this.lblinfo.Text = txtproductcode.Text + "编码检验标准更新成功OK!" + this.txtmodifycheck.Text;
                    this.lblinfo.ForeColor = Color.Blue;
                    txtproductcode.Text = "";
                    txtproductcode.Focus();
                    txtcurrentcheck.Text = "";
                    txtmodifycheck.Text = "";
                }
            }
        }

        private void btnreset_Click(object sender, EventArgs e)
        {
            txtproductcode.Text = "";
            txtproductcode.Focus();
            txtcurrentcheck.Text = "";
            txtmodifycheck.Text = "";
            this.lblinfo.Text = "";
            txtreelid.Text = "";
            txtdatacode.Text = "";
            txtlotno.Text = "";
            txtmaterialcode.Text = "";
            txtqty.Text = "";
            txtvendor.Text = "";
            txtMdate.Text = "";
            txtExpiryDate.Text = "";
        }

        private void txtreelid_Leave(object sender, EventArgs e)
        {
            if (txtreelid.Text.Trim() == "")
                return;
            string sql = "  select reelid,datecode,lotno,materialcode,qty,vendor,Mdate,ExpiryDate from materialRelation  where reelid = '"+txtreelid.Text+"'  order by operdate desc   ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                txtreelid.Text = dt.Rows[0]["reelid"].ToString();
                txtdatacode.Text = dt.Rows[0]["datecode"].ToString();
                txtlotno.Text = dt.Rows[0]["lotno"].ToString();
                txtmaterialcode.Text = dt.Rows[0]["materialcode"].ToString();
                txtqty.Text = dt.Rows[0]["qty"].ToString();
                txtvendor.Text = dt.Rows[0]["vendor"].ToString();
                txtMdate.Text = dt.Rows[0]["Mdate"].ToString();
                txtExpiryDate.Text = dt.Rows[0]["ExpiryDate"].ToString();
            }
        }

        private void txtreelid_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                txtreelid_Leave(null,null);
            }
        }

        private void TestMaterialCheckBySQE_Load(object sender, EventArgs e)
        {
            if (Login.post.Contains("SQE") || Login.manager == "IQC管理员" || Login.manager == "IT管理员")
            {
                btnconfirm.Enabled = true;
            }
            else
            {
                btnconfirm.Enabled = false;
            }

        }
    }
}