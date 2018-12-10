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
    public partial class IQC_batchChage : DevExpress.XtraEditors.XtraForm
    {
        public IQC_batchChage()
        {
            InitializeComponent();
        }

        public IQC_batchChage(DataTable dt)
        {
            InitializeComponent();
            gridControl.DataSource = dt;
        }
        DataTable lotnolist()
        {
            //string sql = "  select reelid 批次号,datecode,convert(varchar(10),Mdate,121) 生产日期,convert(varchar(10),ExpiryDate,121) 有效期 from materialRelation where reelid in ('Z000231640','Z000231641','Z000231642','Z000231643') ";
            string sql = "  select reelid 批次号,datecode,convert(varchar(10),Mdate,121) 生产日期,convert(varchar(10),ExpiryDate,121) 旧有效期,'2019-12-06' 新有效期  from materialRelation where reelid in ('Z000231640','Z000231641')    ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        private void IQC_batchChage_Load(object sender, EventArgs e)
        {           
            //gridControl.DataSource = lotnolist();
        }

        private void Btnconfirm_Click(object sender, EventArgs e)
        {
           
            ArrayList list = new ArrayList();
            list.Clear();

            for (int i = 0; i < gridView.RowCount; i++)
            {
                string lotno =  gridView.GetRowCellValue(i, gridView.Columns["批次号"]).ToString();
                string oldexpiryDate = gridView.GetRowCellValue(i, gridView.Columns["旧有效期"]).ToString();
                string NewexpiryDate = gridView.GetRowCellValue(i, gridView.Columns["新有效期"]).ToString();

                string sql = "";
                sql = "  if  exists( select 1 from IQC_ChageExpiryDate where lotno = '" + lotno + "' )  ";
                sql += "  update IQC_ChageExpiryDate set originalTime='"+oldexpiryDate+ "',itemCounts=itemCounts+1,updateMan= '" + Login.username + "',updateTime=getdate()  where lotno = '" + lotno + "'  ";
                sql += "  else insert into IQC_ChageExpiryDate (lotno ,originalTime,itemCounts,updateMan ,updateTime) values('" +lotno+ "','" +oldexpiryDate+ "',1,'" + Login.username + "', GETDATE())  ";
                sql += "  update materialRelation set ExpiryDate = '"+NewexpiryDate+ "'  where reelid = '" + lotno + "'  ";
                list.Add(sql);           
            }

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
                DialogResult Rt = MessageBox.Show("更改有效期成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (DialogResult.OK == Rt)
                {
                    this.Close();
                }
            }
            else
            {
                MessageBox.Show("更改有效期失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }


    }
}