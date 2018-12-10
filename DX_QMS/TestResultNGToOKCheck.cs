using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using WMSInterface;
using Newtonsoft.Json;
using barcode.WMSInterface;
using DX_QMS.Common;

namespace DX_QMS
{
    public partial class TestResultNGToOKCheck : DevExpress.XtraEditors.XtraForm
    {
        public TestResultNGToOKCheck()
        {
            InitializeComponent();
        }
        public TestResultNGToOKCheck(DataTable dtToERP)
        {
            InitializeComponent();
            databind.DataSource = dtToERP;
        }


        private void UploadToERP(DataGridView db)
        {
            string sqlupdate = "";
            for (int i = 0; i < db.Rows.Count; i++)
            {
                string ss = "True";
                if (db.Rows[i].Cells[0].Value == null)
                    ss = "";
                else
                    ss = db.Rows[i].Cells[0].Value.ToString();

                if (ss == "True")
                {
                    if (int.Parse(db.Rows[i].Cells["qty"].Value.ToString()) < int.Parse(db.Rows[i].Cells["checkqty"].Value.ToString()))
                    {
                        MessageBox.Show("确认数量不能大于可用数量");
                        return;
                    }
                    WMSIQCTestItem wmsiqcTestItem = new WMSIQCTestItem();
                    wmsiqcTestItem.transactionId = db.Rows[i].Cells["iqcTransactionid"].Value.ToString();
                    wmsiqcTestItem.receiptNum = db.Rows[i].Cells["deliveryid"].Value.ToString();
                    wmsiqcTestItem.itemCode = db.Rows[i].Cells["materialcode"].Value.ToString();
                    wmsiqcTestItem.lotNum = db.Rows[i].Cells["lotno"].Value.ToString();
                    wmsiqcTestItem.seqNum = "";
                    wmsiqcTestItem.orgCode = db.Rows[i].Cells["org_id"].Value.ToString();
                    wmsiqcTestItem.iqcResult = db.Rows[i].Cells["TestResult"].Value.ToString() == "OK" ? "S" : "E";
                    wmsiqcTestItem.iqcQty = db.Rows[i].Cells["checkqty"].Value.ToString();

                    List<WMSIQCTestItem> list = new List<WMSIQCTestItem>();
                    list.Add(wmsiqcTestItem);

                    WMSIQCTestInputData wmsIQCTestInputData = new WMSIQCTestInputData();
                    wmsIQCTestInputData.message = list;

                    WMSIQCTestRequestData wmsIQCTestRequestData = new WMSIQCTestRequestData();
                    wmsIQCTestRequestData.iqcRsult = wmsIQCTestInputData;

                    string batchNumn = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                    string method = "CUX_WMS_IQC_REST_R";
                    string requestData = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(wmsIQCTestRequestData)));
                    string responseData = string.Empty;
                    bool restult = WMSInterfaceUtils.CallErpInterface(batchNumn, method, requestData, ref responseData);
                    string s = "0";
                    if (restult)
                        s = "1";
                    else
                        s = "0";
                    sqlupdate = " update delivery set IfSuccess='" + s + "' where lotno='" + db.Rows[i].Cells["lotno"].Value.ToString() + "'";
                    sqlupdate += " insert into WMSInterfaceLog(org_id, workno, batchNumn, method, requestData, result,opertype, uploaddate)";
                    sqlupdate += "values('" + db.Rows[0].Cells["org_id"].Value.ToString() + "','" + db.Rows[0].Cells["deliveryid"].Value.ToString() + "','" + batchNumn + "','" + method + "','" + responseData + "','" + s + "','" + "IQC" + "',getdate())";

                    if (s == "1")
                    {
                        sqlupdate += " if not exists(select 1 from deliveryCheck where lotno='" + db.Rows[i].Cells["lotno"].Value.ToString() + "' and states='" + db.Rows[i].Cells["TestResult"].Value.ToString() + "')";
                        sqlupdate += " insert into deliveryCheck(org_id, receptid,productcode,qty, lotno, states,eventuser, eventtime)";
                        sqlupdate += "values('" + db.Rows[0].Cells["org_id"].Value.ToString() + "','" + db.Rows[0].Cells["deliveryid"].Value.ToString() + "','" + db.Rows[0].Cells["materialcode"].Value.ToString() + "','" + wmsiqcTestItem.iqcQty + "','" + db.Rows[i].Cells["lotno"].Value.ToString() + "','" + db.Rows[i].Cells["TestResult"].Value.ToString() + "','" + Login.userId + "',getdate())";
                        sqlupdate += " else update deliveryCheck set qty=qty+" + wmsiqcTestItem.iqcQty + " where lotno='" + db.Rows[i].Cells["lotno"].Value.ToString() + "' and states='" + db.Rows[i].Cells["TestResult"].Value.ToString() + "'";
                    }
                    DbAccess.ExecuteSql(sqlupdate);
                }
            }
        }

        private void UploadToERP1(DataTable dba)
        {
            WMSIQCTestRequestData wmsIQCTestRequestData = new WMSIQCTestRequestData();
            WMSIQCTestInputData wmsIQCTestInputData = new WMSIQCTestInputData();
            List<WMSIQCTestItem> list = new List<WMSIQCTestItem>();
            string sqlupdate = "";
            for (int k = gridView.GetSelectedRows().Length; k > 0; k--)
            {

                DataRow db = gridView.GetDataRow(gridView.GetSelectedRows()[k - 1]);
                if (db["iqcTransactionid"].ToString() == "")
                    continue;
                if (int.Parse(db["checkqty"].ToString()) <= 0)
                    continue;
                 if (db["TestResult"].ToString() == "")
                    continue;

                if (int.Parse(db["qty"].ToString()) < int.Parse(db["checkqty"].ToString()))
                {
                    MessageBox.Show("确认数量不能大于可用数量");
                    return;
                }
                WMSIQCTestItem wmsiqcTestItem = new WMSIQCTestItem();
                wmsiqcTestItem.transactionId = db["iqcTransactionid"].ToString();
                wmsiqcTestItem.receiptNum = db["deliveryid"].ToString();
                wmsiqcTestItem.itemCode = db["materialcode"].ToString();
                wmsiqcTestItem.lotNum = db["lotno"].ToString();
                wmsiqcTestItem.seqNum = "";
                wmsiqcTestItem.orgCode = db["org_id"].ToString();
                wmsiqcTestItem.iqcResult = db["TestResult"].ToString() == "OK" ? "S" : "E";
                wmsiqcTestItem.iqcQty = db["checkqty"].ToString();
                wmsiqcTestItem.barcodeUser = Login.username;

                sqlupdate += " if not exists(select 1 from deliveryCheck where lotno='" + wmsiqcTestItem.lotNum + "' and states='" + db["TestResult"].ToString() + "')";
                sqlupdate += " insert into deliveryCheck(org_id, receptid,productcode,qty, lotno, states,eventuser, eventtime)";
                sqlupdate += "values('" + wmsiqcTestItem.orgCode + "','" + wmsiqcTestItem.receiptNum + "','" + wmsiqcTestItem.itemCode + "','" + wmsiqcTestItem.iqcQty + "','" + wmsiqcTestItem.lotNum + "','" + db["TestResult"].ToString() + "','" + Login.userId + "',getdate())";
                sqlupdate += " else update deliveryCheck set qty=qty+" + wmsiqcTestItem.iqcQty + " where lotno='" + wmsiqcTestItem.lotNum + "' and states='" + db["TestResult"].ToString() + "'";

                //List<WMSIQCTestItem> list = new List<WMSIQCTestItem>();
                list.Add(wmsiqcTestItem);
                //WMSIQCTestInputData wmsIQCTestInputData = new WMSIQCTestInputData();
                wmsIQCTestInputData.message = list;

            }

            wmsIQCTestRequestData.iqcRsult = wmsIQCTestInputData;
            string batchNumn = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string method = "CUX_WMS_IQC_REST_R";
            string requestData = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(wmsIQCTestRequestData)));
            string responseData = string.Empty;
            bool restult = barcode.WMSInterface.WMSInterfaceUtils.CallErpInterface(batchNumn, method, requestData, ref responseData);
            string s = "";
            if (restult)
                s = "1";
            else
                s = "0";

            if (responseData == "")
            {
                lblinfo.Text = "WMS系统没有返回值,上传失败!";
                return;
            }
            WMSIQCResultHead m = JsonConvert.DeserializeObject<WMSIQCResultHead>(responseData);
            List<WMSIQCInputItem> Items = m.results.messages;
            DataTable dt = new DataTable();
            dt.Columns.Add("transactionId");
            dt.Columns.Add("resultCode");
            dt.Columns.Add("resultMsg");
            string sresult = "";
            foreach (WMSIQCInputItem l in Items)
            {
                if (l.resultCode == "E")
                {
                    dt.Rows.Add(l.transactionId, l.resultCode, l.resultMsg);
                    sresult = sresult + l.transactionId + "," + l.resultMsg + ";";
                }
            }
            for (int j = 0; j < dt.Rows.Count; j++)
            {
                sqlupdate += " delete deliveryCheck where lotno=(select lotno from delivery where iqcTransactionId='" + dt.Rows[j]["transactionId"].ToString() + "')";
            }

            sqlupdate += " insert into WMSInterfaceLog(org_id, workno, batchNumn, method, requestData, result,opertype, uploaddate)";
            sqlupdate += "values('" + dba.Rows[0]["org_id"].ToString() + "','" + dba.Rows[0]["deliveryid"].ToString() + "','" + batchNumn + "','" + method + "','" + responseData + "','" + s + "','" + "IQC" + "',getdate())";

            lblinfo.Text = sresult;

            DbAccess.ExecuteSql(sqlupdate);
        }


        private void btnOK_Click(object sender, EventArgs e)
        {
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }         
            UploadToERP1(dt);
            if (lblinfo.Text != "")
            {
                MessageBox.Show(lblinfo.Text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                DialogResult Rt = MessageBox.Show("全部成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (DialogResult.OK == Rt)
                {
                    this.Close();
                }

            }
            lblinfo.Text = "";
            this.Close();
        }


        private DataTable IQCCheckResultToERP(string lotno)
        {
            //找最后一个批次号对应的接收单号,物料编码的所有检验记录
            string sql = " declare @recept varchar(30),@p varchar(30) ";
            sql += " select @recept=deliveryid,@p=materialcode from delivery where lotno='" + lotno + "' ";
            sql += " SELECT  deliveryid,item, materialcode,d.lotno,d.qty qty,0 OKqty,isnull(finOK.qty,0) NGqty,org_id,iqcTransactionid,'' TestResult,isnull(finOK.qty,0) checkqty  from delivery d inner join ";
            sql += " (select distinct lotno,productcode,receptid,states,isnull(sum(qty),0) qty from deliveryCheck t where receptid=@recept and productcode=@p and states='NG' ";
            sql += "  and not exists(select 1 from deliveryCheck tt where receptid=@recept and productcode=@p and states='OK') ";
            sql += " group by lotno,productcode,receptid,states ";
            sql += " ) finOK on d.lotno=finOK.lotno and d.materialcode=finOK.productcode and d.deliveryid=finOK.receptid ";
            sql += " where deliveryid=@recept and materialcode=@p ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        private void TestResultNGToOKCheck_Load(object sender, EventArgs e)
        {
           /////// databind.DataSource = IQCCheckResultToERP("Z000181628");
        }
    }
}