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
using System.Collections;

namespace DX_QMS.IQCFilePosition
{
    public partial class IQCTestExpiryDate : DevExpress.XtraEditors.XtraForm
    {
        public IQCTestExpiryDate()
        {
            InitializeComponent();
        }


        public IQCTestExpiryDate(string lotno)
        {
            InitializeComponent();
            txtlotno.Text = lotno;
        }
        private void IQCTestExpiryDate_Load(object sender, EventArgs e)
        {
            string sql = "  select ExpiryDate from materialRelation where reelid = '"+txtlotno.Text+ "' ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
            {
                txtoldexpiryDate.Text = dt.Rows[0]["ExpiryDate"].ToString();
            }
        }

        private void txttimeslot_Leave(object sender, EventArgs e)
        {

            int output = 0;
            if (!int.TryParse(txttimeslot.Text.Trim(), out output))
            {
                MessageBox.Show("时间段格式不正确","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                txttimeslot.Text = "";
                txttimeslot.Focus();
                return;
            }

        }

        private void sBtnOK_Click(object sender, EventArgs e)
        {
            if (txtlotno.Text == "" || txtoldexpiryDate.Text=="" || txttimeslot.Text == "")
            {
                MessageBox.Show("请输入完整的信息", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            int delaydays = 0;
            if (!int.TryParse(txttimeslot.Text.Trim(), out delaydays))
            {
                MessageBox.Show("时间段格式不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txttimeslot.Text = "";
                txttimeslot.Focus();
                return;
            }

            ArrayList list = new ArrayList();
            list.Clear();

            string sql = "  if  exists( select 1 from IQC_ChageExpiryDate where lotno = '"+txtlotno.Text+ "' )  ";
                   sql += "  update IQC_ChageExpiryDate set delayDays="+delaydays+ " ,itemCounts=itemCounts+1,updateMan= '" + Login.username+ "',updateTime=getdate()  where lotno = '"+txtlotno.Text+ "'  ";
                   sql += "  else insert into IQC_ChageExpiryDate (lotno ,originalTime ,delayDays,itemCounts,updateMan ,updateTime) values('" + txtlotno.Text+ "','"+txtoldexpiryDate.Text+ "',"+ delaydays+ ",1,'"+Login.username+ "', GETDATE())  ";
            list.Add(sql);

            sql = "  update materialRelation set ExpiryDate = dateadd(day,"+delaydays+ ",ExpiryDate)  where reelid = '" + txtlotno.Text + "'  ";
            list.Add(sql);

            bool flag = false;
            try
            {
                flag = DbAccess.ExecutSqlTran(list);
            }
            catch
            {
                MessageBox.Show("更改有效期失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (flag == true)
            {
                MessageBox.Show("更改有效期成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);          
            }
            else
            {
                MessageBox.Show("更改有效期失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void sBtnreset_Click(object sender, EventArgs e)
        {
            txttimeslot.Text = "";
        }
    }
}