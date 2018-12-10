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

namespace DX_QMS
{
    public partial class TestSampleList : DevExpress.XtraEditors.XtraForm
    {
        public TestSampleList()
        {
            InitializeComponent();
        }
        public TestSampleList(string sampletype, string productcode, string supp)
        {
            InitializeComponent();
            txtsampletype.Text = sampletype;
            txtproductcode.Text = productcode;
            txtsupp.Text = supp;
            BindTestSample(sampletype, productcode, supp);
        }
        private void BindTestSample(string sampletype, string productcode, string supp)
        {
            string sql = "select sampletype, productcode, item, qty, operuser, operdate from IQC_SampleList where sampletype='" + sampletype + "' and productcode='" + productcode + "' and supplier='" + supp + "'";
            DataSet ds = Common.DbAccess.SelectBySql(sql);
            databind.DataSource = ds.Tables[0];
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {



            string sql = "declare @timekey varchar(50) select @timekey = CONVERT(varchar(30),getdate(), 121) ";
            sql += " insert into IQC_SampleList(sampletype,supplier,productcode, item, qty, operuser, operdate)";
            sql += " values('" + this.txtsampletype.Text + "','" + txtsupp.Text + "','" + txtproductcode.Text + "',@timekey,'" + this.txtqty.Text + "','" + Login.username + "',getdate()" + ")";
            if (Common.DbAccess.ExecuteSql(sql))
            {
                BindTestSample(txtsampletype.Text, txtproductcode.Text, txtsupp.Text);
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            DataTable dt = databind.DataSource as DataTable;
            if (dt==null || dt.Rows .Count <0)
            {
                return;
            }
            if (gridView.RowCount <= 0)
                return;
            else
            {
                string sql = "delete from IQC_SampleList where productcode='" + gridView.GetFocusedRowCellValue("productcode").ToString() + "' and item='" + gridView.GetFocusedRowCellValue("item").ToString() + "'";
                if (Common.DbAccess.ExecuteSql(sql))
                {
                    MessageBox.Show("删除成功");
                }
            }
            BindTestSample(txtsampletype.Text, txtproductcode.Text, txtsupp.Text);
        }
    }
}