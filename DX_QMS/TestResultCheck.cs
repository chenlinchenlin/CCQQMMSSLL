using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using WMSInterface;
using Newtonsoft.Json;
using DX_QMS.Common;
using barcode.WMSInterface;
using DevExpress.XtraBars;
using System.Data.SqlClient;

namespace DX_QMS
{
    public partial class TestResultCheck : DevExpress.XtraEditors.XtraForm
    {
        int StockQty =0;
        string  states = "OK";
        string Mouldmaterialcode = "";
        string sMaterialname = "";
        string WMSremark = "";

        public TestResultCheck()
        {
            InitializeComponent();
        }
        public TestResultCheck( DataTable dtToERP)
        {
            InitializeComponent();
            databind.DataSource = dtToERP;          
        }


        public TestResultCheck(int stockqty, string States, string remark,string Materialname, DataTable dtToERP)
        {
            InitializeComponent();

            StockQty = stockqty;
            states = States;
            WMSremark = remark;
            sMaterialname = Materialname;
            databind.DataSource = dtToERP;
        }


        private void UploadToERP(DataTable dba)
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

            string responseoriginalDate = string.Empty;
            //bool restult = barcode.WMSInterface.WMSInterfaceUtils.CallWMSInterface(batchNumn, method, requestData, ref responseData, ref responseoriginalDate);

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

            sqlupdate += " insert into WMSInterfaceLog(org_id, workno, batchNumn, method, requestData, result,opertype, uploaddate,originalrequest,originalresponseDate)";
            sqlupdate += "values('" + dba.Rows[0]["org_id"].ToString() + "','" + dba.Rows[0]["deliveryid"].ToString() + "','" + batchNumn + "','" + method + "','" + responseData + "','" + s + "','" + "IQC" + "',getdate(),'"+ requestData +"','"+ responseoriginalDate +"')";

            lblinfo.Text = sresult;

            DbAccess.ExecuteSql(sqlupdate);
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

                if (int.Parse(db["qty"].ToString()) < int.Parse(db["checkqty"].ToString()))
                {
                    MessageBox.Show("确认数量不能大于可用数量");
                    return;
                }

                if (db["TestResult"].ToString() == "")
                {
                    continue;
                }
                WMSIQCTestItem wmsiqcTestItem = new WMSIQCTestItem();
                wmsiqcTestItem.transactionId = db["iqcTransactionid"].ToString();
                wmsiqcTestItem.comments = db["remark"].ToString();
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

                Mouldmaterialcode = db["materialcode"].ToString();

                //List<WMSIQCTestItem> list = new List<WMSIQCTestItem>();
                list.Add(wmsiqcTestItem);
                //WMSIQCTestInputData wmsIQCTestInputData = new WMSIQCTestInputData();
                wmsIQCTestInputData.message = list;


                /*
                if (!sMaterialname.Contains("【F1"))
                {
                    string connPANASql = "server=10.100.64.4,6815;database=PanaCIM;user id=sa;password=PANASONIC1!;";
                    SqlConnection sqlconn = new SqlConnection(connPANASql);
                    if (sqlconn.State == ConnectionState.Closed)
                        sqlconn.Open();
                    SqlTransaction tran = sqlconn.BeginTransaction();
                    DataTable Uploaddt = GetLotList(db["lotno"].ToString().Trim());
                    if (Uploaddt.Rows.Count > 0)
                    {
                        string result = WriteDataToPANASON(Uploaddt.Rows[0]["reelid"].ToString(), Uploaddt.Rows[0]["materialcode"].ToString(), Uploaddt.Rows[0]["Vendor"].ToString(), Uploaddt.Rows[0]["cuslotno"].ToString(), int.Parse(Uploaddt.Rows[0]["qty"].ToString()), Uploaddt.Rows[0]["datecode"].ToString(), sqlconn, tran);
                        if (result == "0")
                        {
                            tran.Commit();
                            //this.lblMessage.Text = Uploaddt.Rows[0]["reelid"].ToString() + "保存成功!" + ",写入松下数据成功!";
                        }
                        else if (result == "1")
                        {
                           // this.lblMessage.Text = "保存成功!" + ",松下系统,已经存在该数据!" + Uploaddt.Rows[0]["reelid"].ToString();
                            tran.Rollback();
                        }
                        else
                        {
                            //this.lblMessage.Text = Uploaddt.Rows[0]["reelid"].ToString() + "保存成功!" + ",写入松下数据失败!" + result + ",请重试一次";
                            // this.lblMessage.ForeColor = Color.Red;
                            //Component.UrlEncoding.alertMess(ClientScript, lblMessage.Text);
                            tran.Rollback();

                            string sqlinsert = " if not exists(select 1 from OEM_ReelidDataUpload where Lotno='" + Uploaddt.Rows[0]["reelid"].ToString() + "') insert into OEM_ReelidDataUpload(Lotno, materialcode, vendor, cuslotno,qty, datecode,operdate) ";
                            sqlinsert += " values('" + Uploaddt.Rows[0]["reelid"].ToString() + "','" + Uploaddt.Rows[0]["materialcode"].ToString() + "','" + Uploaddt.Rows[0]["Vendor"].ToString() + "','" + Uploaddt.Rows[0]["cuslotno"].ToString() + "','" + Uploaddt.Rows[0]["qty"].ToString() + "','" + Uploaddt.Rows[0]["datecode"].ToString() + "',getdate())";
                            DbAccess.ExecuteSql(sqlinsert);
                        }
                    }
                 }
                */

            }

            wmsIQCTestRequestData.iqcRsult = wmsIQCTestInputData;
            string batchNumn = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string method = "CUX_WMS_IQC_REST_R";
            string requestData = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(wmsIQCTestRequestData)));
            string responseData = string.Empty;

            bool restult = barcode.WMSInterface.WMSInterfaceUtils.CallErpInterface(batchNumn, method, requestData, ref responseData);

            ////string responseoriginalDate = string.Empty;
            ////bool restult = barcode.WMSInterface.WMSInterfaceUtils.CallWMSInterface(batchNumn, method, requestData, ref responseData, ref responseoriginalDate);

            string s = "";
            if (restult)
                s = "1";
            else
                s = "0";

            if (responseData == "")
            {
                lblinfo.Text = "WMS系统没有返回报文,入库失败!";
                return;
            }
            WMSIQCResultHead m = JsonConvert.DeserializeObject<WMSIQCResultHead>(responseData);
            List<WMSIQCInputItem> Items = m.results.messages;
            DataTable dt = new DataTable();
            dt.Columns.Add("transactionId");
            dt.Columns.Add("resultCode");
            dt.Columns.Add("resultMsg");
            string sresult = "";
            string responseDatasql = "";
            foreach (WMSIQCInputItem l in Items)
            {


                //responseDatasql += " if not exists(select 1 from WMS_responseData where transactionId='" + l.transactionId + "')";
                //responseDatasql += "  insert into WMS_responseData ( transactionId ,resultCode,resultMsg,updateTime ) ";
                //responseDatasql += "  values ('" + l.transactionId + "','" + l.resultCode + "','" + l.resultMsg + "',GETDATE ())  ";
                //responseDatasql += " else update WMS_responseData set resultCode='" + l.resultCode + "',resultMsg  = '" + l.resultMsg + "',updateTime=GETDATE ()  where transactionId='" + l.transactionId + "'";


                if ((l.resultCode == "E" && !(l.resultMsg.Contains("该笔数据已经同步过合格IQC结果，不允许重复。"))) || l.resultCode == "N")     // 还有一种 "N"
                {
                    dt.Rows.Add(l.transactionId, l.resultCode, l.resultMsg);
                    sresult = sresult + l.transactionId + "," + l.resultMsg + ";";
                }

                //if ((! l.resultMsg.Contains("同步成功。")) || (!l.resultMsg.Contains("该笔数据已经同步过合格IQC结果")))
                //{
                //    dt.Rows.Add(l.transactionId, l.resultCode, l.resultMsg);
                //    sresult = sresult + l.transactionId + "," + l.resultMsg + ";";
                //}

            }
            for (int j = 0; j < dt.Rows.Count; j++)
            {
                sqlupdate += " delete deliveryCheck where lotno=(select lotno from delivery where iqcTransactionId='" + dt.Rows[j]["transactionId"].ToString() + "')";
            }

            sqlupdate += " insert into WMSInterfaceLog(org_id, workno, batchNumn, method, requestData, result,opertype, uploaddate)";
            sqlupdate += "values('" + dba.Rows[0]["org_id"].ToString() + "','" + dba.Rows[0]["deliveryid"].ToString() + "','" + batchNumn + "','" + method + "','" + responseData + "','" + s + "','" + "IQC" + "',getdate())";
    

            ////sqlupdate += " insert into WMSInterfaceLog(org_id, workno, batchNumn, method, requestData, result,opertype, uploaddate,originalrequest,originalresponseDate)";
            ////sqlupdate += "values('" + dba.Rows[0]["org_id"].ToString() + "','" + dba.Rows[0]["deliveryid"].ToString() + "','" + batchNumn + "','" + method + "','" + responseData + "','" + s + "','" + "IQC" + "',getdate(),'" + requestData + "','" + responseoriginalDate + "')";



            lblinfo.Text = sresult;

            DbAccess.ExecuteSql(sqlupdate);
            //DbAccess.ExecuteSql(responseDatasql);

        }

        bool updateMould(string productcode, string thisTimeQty)
        {
            string sql = "  if  exists(select 1 from IQC_MouldCycle where materialcode = '" + productcode + "' )   ";
            sql += "  update IQC_MouldCycle set RestQty = RestQty - " + thisTimeQty + " ,thisTimeQty = " + thisTimeQty + " ,updateMan = '" + Login.username + "' ,updateTime = GETDATE() ";
            sql += " where materialcode ='" + productcode + "'";
            return DbAccess.ExecuteSql(sql);
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                MessageBox.Show("没有数据","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }
            UploadToERP1(dt);
            try
            {
                updateMould(Mouldmaterialcode, StockQty.ToString());
            }
            catch {
            }
            if (lblinfo.Text != "")
            {
                MessageBox.Show(lblinfo.Text, "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
               // MessageBox.Show("请确认所选的批次号都有接收单号", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                //MessageBox.Show("全部OK", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult Rt = MessageBox.Show("数据推送成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (DialogResult.OK == Rt)
                {
                    this.Close();
                }

            }
            lblinfo.Text = "";
            this.Close();
        }

        private void cbcheck_CheckedChanged(object sender, EventArgs e)
        {

            //for (int i = 0; i < gridView.RowCount; i++)
            //{
            //    //databind.Rows[i].Cells[0].Value = cbcheck.Checked;
            //    gridView.SetRowCellValue(i, gridView.Columns[0], cbcheck.Checked);
            //}

        }


        private DataTable GetLotList(string lotno)
        {
            string sql = "select deliveryid,d.materialcode,materialname,r.lotno cuslotno,reelid,d.lotno,r.qty,replace(vendor,'&','') vendor,replace(mfr,'&','') mfr,vendorname,dateadd(day,isnull(usefullife,90),isnull(Mdate,transactiondate))  ExpiryDate, ";
            sql += " transactiondate,replace(datecode,'&','') datecode,r.materialcode cuscode,isnull(Mdate,transactiondate) Mdate from materialRelation r inner join delivery d on r.reelid=d.lotno left join MaterialSpec m on  d.materialcode=m.materialcode where r.reelid='" + lotno + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        public string WriteDataToPANASON(string ReelBarcode, string PartNumber, string Vendor, string Lot, int InitialQuantity, string UserData, SqlConnection cn, SqlTransaction tran)
        {

            SqlCommand sqlCmd;
            SqlParameter sqlParm;

            sqlCmd = new SqlCommand();
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.CommandText = "Insert_datatoreel";
            sqlCmd.Connection = cn;
            sqlCmd.Transaction = tran;

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@reelid";
            sqlParm.SqlDbType = SqlDbType.NVarChar;
            sqlParm.Value = ReelBarcode;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@parcode";
            sqlParm.SqlDbType = SqlDbType.NVarChar;
            sqlParm.Value = PartNumber;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@vendorcode";
            sqlParm.SqlDbType = SqlDbType.NVarChar;
            sqlParm.Value = Vendor;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@lotno";
            sqlParm.SqlDbType = SqlDbType.NVarChar;
            sqlParm.Value = Lot;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@qty";
            sqlParm.SqlDbType = SqlDbType.Int;
            sqlParm.Value = InitialQuantity;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@datecode";
            sqlParm.SqlDbType = SqlDbType.NVarChar;
            sqlParm.Value = UserData;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter("@errormsg", SqlDbType.VarChar, 20);
            sqlParm.Direction = ParameterDirection.Output;
            sqlCmd.Parameters.Add(sqlParm);

            sqlCmd.ExecuteNonQuery();

            return sqlCmd.Parameters["@errormsg"].Value.ToString();
        }


        private DataTable IQCCheckResultToERP(string lotno)
        {
            string sql = " declare @recept varchar(30),@p varchar(30)";
            sql += "select @recept=receptid,@p=Productcode from IQC_TestList where LotNo='" + lotno + "'";
            sql += " SELECT  deliveryid,item, materialcode,d.lotno,(d.qty-isnull(finOK.qty,0)-isnull(finNG.qty,0)) qty,finOK.qty OKqty,finNG.qty NGqty,org_id,iqcTransactionid,'' TestResult,(d.qty-isnull(finOK.qty,0)-isnull(finNG.qty,0)) checkqty ,'' remark  from delivery d left join";
            sql += "(select distinct lotno,productcode,receptid,states,isnull(sum(qty),0) qty from deliveryCheck t where receptid=@recept and productcode=@p and states='OK' group by lotno,productcode,receptid,states";
            sql += ") finOK on d.lotno=finOK.lotno and d.materialcode=finOK.productcode and d.deliveryid=finOK.receptid  left join ";
            sql += "(select lotno,productcode,receptid,states,isnull(sum(qty),0) qty from deliveryCheck t where receptid=@recept and productcode=@p and states='NG' group by lotno,productcode,receptid,states";
            sql += ") finNG on d.lotno=finNG.lotno and d.materialcode=finNG.productcode and d.deliveryid=finNG.receptid  ";
            sql += " where deliveryid=@recept and materialcode=@p and org_id = 'HCL' ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        private void TestResultCheck_Load(object sender, EventArgs e)
        {
            lblinfo.Text = "请选择有接收单号的批次号";
            ///// databind.DataSource = IQCCheckResultToERP("Z000197162");


            //cbok.Enabled = false;
            //cbNG.Enabled = false;

            int CountQty = 0, sstockqty = 0;   ///// StockQty
            sstockqty = StockQty;

            for (int i = 0; i < gridView.RowCount; i++)
            {
                if (int.Parse(gridView.GetRowCellValue(i, gridView.Columns["checkqty"]).ToString()) > 0)
                {
                    if (sstockqty <= int.Parse(gridView.GetRowCellValue(i, gridView.Columns["checkqty"]).ToString()))
                    {
                        gridView.SelectRow(i);
                        gridView.SetRowCellValue(i, gridView.Columns["checkqty"], sstockqty.ToString());
                        gridView.SetRowCellValue(i, gridView.Columns["TestResult"], states);
                        break;
                    }

                    CountQty += int.Parse(gridView.GetRowCellValue(i, gridView.Columns["checkqty"]).ToString());
                    sstockqty -= int.Parse(gridView.GetRowCellValue(i, gridView.Columns["checkqty"]).ToString());

                    gridView.SelectRow(i);
                    gridView.SetRowCellValue(i, gridView.Columns["TestResult"], states);


                    //break;
                }

                if (i < 20)
                {
                    gridView.SetRowCellValue(i, gridView.Columns["remark"], WMSremark);
                }
            }
        }

        //private void cbok_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (cbok.Checked)
        //    {
        //        for (int i = 0; i < gridView.RowCount; i++)
        //        {
                 
        //            gridView.SetRowCellValue(i, gridView.Columns["TestResult"], "OK");

        //        }
        //        cbNG.Checked = false;
        //    }
        //}

        //private void cbNG_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (cbNG.Checked)
        //    {
        //        for (int i = 0; i < gridView.RowCount; i++)
        //        {
        //            gridView.SetRowCellValue(i, gridView.Columns["TestResult"], "NG");
        //        }
        //        cbok.Checked = false;
        //    }
        //}

        private void txtlotno_KeyUp(object sender, KeyEventArgs e)
        {

            if (e.KeyValue == 13 && txtlotno.Text != "")
            {
                int k = 0;
                for (int i = 0; i < gridView.RowCount; i++)
                {
                    if (txtlotno.Text.Trim().ToUpper() == gridView.GetRowCellValue(i, gridView.Columns["lotno"]).ToString())
                    {
                        gridView.SelectRow(i);         
                        gridView.SetRowCellValue(i, gridView.Columns["TestResult"], "OK");
                        txtlotno.Text = "";
                        txtlotno.Focus();
                        break;
                    }
                    else
                        k++;
                }
                if (k == gridView.RowCount)
                {
                    MessageBox.Show(txtlotno.Text + ",该批次号不是该接收单号,请重新扫描!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtlotno.Text = "";
                    txtlotno.Focus();
                }
            }

        }


        private void txtproductcode_KeyUp(object sender, KeyEventArgs e)
        {

            if (e.KeyValue == 13 && txtproductcode.Text != "")
            {
                int k = 0;
                for (int i = 0; i < gridView.RowCount; i++)
                {
                    if (txtproductcode.Text.Trim().ToUpper() == gridView.GetRowCellValue(i, gridView.Columns["materialcode"]).ToString())
                    {
                        //gridView.SelectRow(i);          
                        //gridView.SetRowCellValue(i, gridView.Columns["TestResult"], "OK");

                        labmessage.Text = "物料编码正确！";
                        labmessage.ForeColor = Color.Blue;

                        txtproductcode.Text = "";
                        txtproductcode.Focus();
                        break;
                    }
                    else
                        k++;
                }
                if (k == gridView.RowCount)
                {
                    labmessage.Text = "物料编码错误！";
                    labmessage.ForeColor = Color.Red;
                    MessageBox.Show(txtproductcode.Text + ",物料编码不是该接收单号,请认真核对!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtproductcode.Text = "";
                    txtproductcode.Focus();
                }
            }

        }
        private void gridView_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            gridView.IndicatorWidth = 50;

            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }
    }
}