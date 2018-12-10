using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DX_QMS.Common;

namespace DX_QMS
{
    public partial class CheckEMSCusLotno : DevExpress.XtraEditors.XtraForm
    {
        DataTable dtems = new DataTable();
        string stockstate = "OK";
        string stocklotno = "";
        string stockremark = "";
        int countcuscode = 0;
        string materialcode = "";
        string sMaterialname = "";

        public CheckEMSCusLotno()
        {
            InitializeComponent();
        }

        public CheckEMSCusLotno( int stockqty,string stestresult,string slotno,string sremark,string Materialname, DataTable dtEMS, DataTable DtCheck)
        {
            InitializeComponent();
            dtems = dtEMS;

            txtstockqty.Text = stockqty.ToString();

            stockstate = stestresult;
            stocklotno = slotno;
            stockremark = sremark;
            sMaterialname = Materialname;
            gridControl.DataSource = DtCheck;
        }


        private DataTable IQCCheckEMSCusLotnoResult(string lotno)
        {
            string sql = " declare @recept varchar(30),@p varchar(30)";
            sql += "select @recept=receptid,@p=Productcode from IQC_TestList where LotNo='" + lotno + "'";
            sql += " SELECT  deliveryid,item, materialcode,d.lotno,(d.qty-isnull(finOK.qty,0)-isnull(finNG.qty,0)) qty,finOK.qty OKqty,finNG.qty NGqty,org_id,iqcTransactionid,'' TestResult,(d.qty-isnull(finOK.qty,0)-isnull(finNG.qty,0)) checkqty ,'' remark from delivery d left join";
            sql += "(select distinct lotno,productcode,receptid,states,isnull(sum(qty),0) qty from deliveryEMSCusCheck t where receptid=@recept and productcode=@p and states='OK' group by lotno,productcode,receptid,states";
            sql += ") finOK on d.lotno=finOK.lotno and d.materialcode=finOK.productcode and d.deliveryid=finOK.receptid  left join ";
            sql += "(select lotno,productcode,receptid,states,isnull(sum(qty),0) qty from deliveryEMSCusCheck t where receptid=@recept and productcode=@p and states='NG' group by lotno,productcode,receptid,states";
            sql += ") finNG on d.lotno=finNG.lotno and d.materialcode=finNG.productcode and d.deliveryid=finNG.receptid  ";
            sql += " where deliveryid=@recept and materialcode=@p ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }


        private DataTable IQCCheckEMSCusLotno(string lotno)
        {
            string sql = " declare @recept varchar(30),@p varchar(30)";
            sql += "select @recept=receptid,@p=Productcode from IQC_TestList where LotNo='" + lotno + "'";
            sql += "  SELECT  deliveryid,item, materialcode,lotno,qty,vendorcode,vendorname,org_id,iqcTransactionid,'' TestResult,'' remark from delivery ";
            sql += " where deliveryid=@recept and materialcode=@p ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        private void CheckEMSCusLotno_Load(object sender, EventArgs e)
        {
           //////gridControl.DataSource = IQCCheckEMSCusLotno("L16A8000942");
           labmessage.Text = " 至少输入一个客户/供应商料号！";

        }




        public string insertBarCode(string org_id, string item_id, string tran_id, Int32 pqty, Int32 enableqty, string baruser, string receptid, string productcode, string lotno, string states, string remarks, SqlConnection cn, SqlTransaction tran)
        {

            SqlCommand sqlCmd;
            SqlParameter sqlParm;

            sqlCmd = new SqlCommand();
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.CommandText = "Warehouse_IQCCheckInputAdd";
            sqlCmd.Connection = cn;
            sqlCmd.Transaction = tran;

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@org_id";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = org_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@item_id";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = item_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@tran_id";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = tran_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@pqty";
            sqlParm.SqlDbType = SqlDbType.Int;
            sqlParm.Value = pqty;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@enableqty";
            sqlParm.SqlDbType = SqlDbType.Int;
            sqlParm.Value = enableqty;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@userid";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = baruser;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@receptid";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = receptid;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@productcode";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = productcode;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@lotno";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = lotno;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@states";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = states;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "@remarks";
            sqlParm.SqlDbType = SqlDbType.VarChar;
            sqlParm.Value = remarks;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            sqlParm.Direction = ParameterDirection.Output;
            sqlCmd.Parameters.Add(sqlParm);

            sqlCmd.ExecuteNonQuery();
            return sqlCmd.Parameters["@msg"].Value.ToString();

        }



        public string IQC_TestOtherInstock(string org_id, string item_id, string tran_id, string receptid, string productcode, string pqty, string enableqty, string userid, string lotno, string states, string remarks)
        {

            SqlParameter[] para = new SqlParameter[12];
            para[0] = new SqlParameter("@org_id", org_id);
            para[1] = new SqlParameter("@item_id", item_id);
            para[2] = new SqlParameter("@tran_id", tran_id);
            para[3] = new SqlParameter("@receptid", receptid);
            para[4] = new SqlParameter("@productcode", productcode);
            para[5] = new SqlParameter("@pqty", pqty);
            para[6] = new SqlParameter("@enableqty", enableqty);
            para[7] = new SqlParameter("@userid", userid);
            para[8] = new SqlParameter("@lotno", lotno);
            para[9] = new SqlParameter("@states", states);
            para[10] = new SqlParameter("@remarks", remarks);
            para[11] = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            para[11].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Warehouse_IQCCheckInputAdd", para);
            return para[11].Value.ToString();
        }

        protected int OtherReceiveInstock(string org_id, string sDelivery, string productcode)
        {
            int qty = 0;
            string ssql = "select isnull(sum(qty),0) qty from Warehouse_IQCCheck where org_id='" + org_id + "' and receptid='" + sDelivery + "' and productcode='" + productcode + "'";
            DataTable dt = Common.DbAccess.SelectBySql(ssql).Tables[0];
            if (dt.Rows.Count > 0)
                qty = int.Parse(dt.Rows[0]["qty"].ToString());
            return qty;
        }

        bool updateMould(string productcode, string thisTimeQty)
        {
            string sql = "  if  exists(select 1 from IQC_MouldCycle where materialcode = '"+ productcode + "' )   ";
            sql += "  update IQC_MouldCycle set RestQty = RestQty - "+ thisTimeQty +" ,thisTimeQty = "+ thisTimeQty+" ,updateMan = '" + Login.username + "' ,updateTime = GETDATE() ";
            sql += " where materialcode ='" + productcode + "'";
            return DbAccess.ExecuteSql(sql);
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

        private DataTable GetLotList(string lotno)
        {
            string sql = "select deliveryid,d.materialcode,materialname,r.lotno cuslotno,reelid,d.lotno,r.qty,replace(vendor,'&','') vendor,replace(mfr,'&','') mfr,vendorname,dateadd(day,isnull(usefullife,90),isnull(Mdate,transactiondate))                 ExpiryDate, ";
            sql += " transactiondate,replace(datecode,'&','') datecode,r.materialcode cuscode,isnull(Mdate,transactiondate) Mdate from materialRelation r inner join delivery d on r.reelid=d.lotno left join MaterialSpec m on         d.materialcode=m.materialcode where r.reelid='" + lotno + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }



        private void btnOK_Click(object sender, EventArgs e)
        {

            DataTable checklotno = gridControl.DataSource as DataTable;
            if (checklotno == null || checklotno.Rows.Count < 1)
            {
                return;
            }

              //if (gridView.GetSelectedRows().Length < 1)
             //    return;


            if (countcuscode < 1)
            {
                MessageBox.Show("至少需要扫一次客户/供应商料号", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            for (int n = 0; n < checklotno.Rows.Count; n++)
            {
                if (!sMaterialname.Contains("【F1"))
                {
                    string connPANASql = "server=10.100.64.4,6815;database=PanaCIM;user id=sa;password=PANASONIC1!;";
                    SqlConnection sqlconn = new SqlConnection(connPANASql);
                    if (sqlconn.State == ConnectionState.Closed)
                        sqlconn.Open();
                    SqlTransaction tran = sqlconn.BeginTransaction();
                    DataTable Uploaddt = GetLotList(checklotno.Rows[n]["lotno"].ToString());
                    if (Uploaddt.Rows.Count > 0)
                    {
                        string result = WriteDataToPANASON(Uploaddt.Rows[0]["reelid"].ToString(), Uploaddt.Rows[0]["materialcode"].ToString(), Uploaddt.Rows[0]["Vendor"].ToString(), Uploaddt.Rows[0]["cuslotno"].ToString(), int.Parse(Uploaddt.Rows[0]["qty"].ToString()), Uploaddt.Rows[0]["datecode"].ToString(), sqlconn, tran);
                        if (result == "0")
                        {
                            tran.Commit();
                            // this.lblMessage.Text = Uploaddt.Rows[0]["reelid"].ToString() + "保存成功!" + ",写入松下数据成功!";
                        }
                        else if (result == "1")
                        {
                            //this.lblMessage.Text = "保存成功!" + ",松下系统,已经存在该数据!" + Uploaddt.Rows[0]["reelid"].ToString();
                            tran.Rollback();
                        }
                        else
                        {
                            //this.lblMessage.Text = Uploaddt.Rows[0]["reelid"].ToString() + "保存成功!" + ",写入松下数据失败!" + result + ",请重试一次";
                            // this.lblMessage.ForeColor = Color.Red;
                            // Component.UrlEncoding.alertMess(ClientScript, lblMessage.Text);
                            tran.Rollback();
                            string sqlinsert = " if not exists(select 1 from OEM_ReelidDataUpload where Lotno='" + Uploaddt.Rows[0]["reelid"].ToString() + "') insert into OEM_ReelidDataUpload(Lotno, materialcode, vendor, cuslotno,qty, datecode,operdate) ";
                            sqlinsert += " values('" + Uploaddt.Rows[0]["reelid"].ToString() + "','" + Uploaddt.Rows[0]["materialcode"].ToString() + "','" + Uploaddt.Rows[0]["Vendor"].ToString() + "','" + Uploaddt.Rows[0]["cuslotno"].ToString() + "','" + Uploaddt.Rows[0]["qty"].ToString() + "','" + Uploaddt.Rows[0]["datecode"].ToString() + "',getdate())";
                            DbAccess.ExecuteSql(sqlinsert);
                        }
                    }
                }
            }

            //int OKqty = 0, NGqty = 0;
            //string sql = "";
            //string OKLotno = "", NGLotno = "";
            //string OKremark = "", NGremark = "";

            //for (int k = gridView.GetSelectedRows().Length; k > 0; k--)
            //{
            //    DataRow db = gridView.GetDataRow(gridView.GetSelectedRows()[k - 1]);

            //    if (db["TestResult"].ToString() == "")
            //    {
            //        continue;
            //    }

            //if (int.Parse(db["qty"].ToString()) < int.Parse(db["checkqty"].ToString()))
            //{
            //    MessageBox.Show("确认数量不能大于可用数量");
            //    return;
            //}

            //if (db["TestResult"].ToString() == "OK" && int.Parse(db["qty"].ToString()) > 0)
            //{
            //    OKqty = OKqty + int.Parse(db["checkqty"].ToString());
            //    OKLotno = db["lotno"].ToString();
            //    OKremark = db["remark"].ToString();

            //}
            //if (db["TestResult"].ToString() == "NG" && int.Parse(db["qty"].ToString()) > 0)
            //{
            //    NGqty = NGqty + int.Parse(db["checkqty"].ToString());
            //    NGLotno = db["lotno"].ToString();
            //    NGremark = db["remark"].ToString();
            //}

            //sql += " if not exists(select 1 from deliveryEMSCusCheck where lotno='" + db["lotno"].ToString() + "' and states='" + db["TestResult"].ToString() + "')";
            //sql += " insert into deliveryEMSCusCheck(org_id, receptid,productcode,qty, lotno, states,eventuser, eventtime)";
            //sql += "values('" + db["org_id"].ToString() + "','" + db["deliveryid"].ToString() + "','" + db["materialcode"].ToString() + "','" + db["checkqty"].ToString() + "','" + db["lotno"].ToString() + "','" + db["TestResult"].ToString() + "','" + Login.userId + "',getdate())";
            //sql += " else update deliveryEMSCusCheck set qty=qty+" + db["checkqty"].ToString() + " where lotno='" + db["lotno"].ToString() + "' and states='" + db["TestResult"].ToString() + "'";

            // }

            SqlConnection conn = new SqlConnection(DbAccess.connSql);
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            SqlTransaction tran1 = conn.BeginTransaction();

            try
           {


                //if (OKqty >= 0)
                //{
                string states = "OK";
                if (stockstate == "OK")
                    states = "OK";
                else
                    states = "NG";

                int stockqty = 0;
                if (txtstockqty.Text != "")
                    stockqty = int.Parse(txtstockqty.Text);
                else
                    return;

                string lotno = "";
                if (stocklotno != "")
                    lotno = stocklotno;
                else
                    return;

                string remarks = "";
                if (stockremark != "")
                    remarks = stockremark;

                string org_id = dtems.Rows[0]["org_id"].ToString();
                string srecept = dtems.Rows[0]["deliveryid"].ToString();
                string productcode = dtems.Rows[0]["materialcode"].ToString();
                string item_id = dtems.Rows[0]["INVENTORY_ITEM_ID"].ToString();
                int alreadyqty = OtherReceiveInstock(org_id, srecept, productcode);
                string enableqty = dtems.Rows[0]["qty"].ToString();
                ////string EMSCusSql = @" delete Warehouse_IQCCheck where org_id='"+ org_id + "' and transaction_id = '"+ productcode + "' and  item = '"++"' ";

                //还可入库数量(大于将要入库数量)
                if (int.Parse(enableqty == "" ? "0" : enableqty) - alreadyqty >= stockqty)
                {
                        //  string barcodemsg = IQC_TestOtherInstock(org_id, item_id, productcode, srecept, productcode, stockqty.ToString(), enableqty, Login.username, lotno, states, remarks);

                        string barcodemsg =  insertBarCode(org_id, item_id, productcode, stockqty, Int32.Parse(enableqty),Login.username, srecept, productcode, lotno, states, remarks, conn, tran1);

                    if (barcodemsg.IndexOf("成功") >= 0)
                    {
                        //bool flag = DbAccess.ExecuteSql(sql);
                        //if (flag == false)
                        //{
                        //    lblinfoinstock.ForeColor = Color.Red;
                        //    lblinfoinstock.Text = "入库失败";
                        //    tran1.Rollback();
                        //    return;
                        //}
                        try
                        {
                            updateMould(productcode,txtstockqty.Text.Trim());
                        }
                        catch
                        {
                        }
                        tran1.Commit();
                        lblinfoinstock.ForeColor = Color.Blue;
                        lblinfoinstock.Text = barcodemsg;
                        DialogResult Rt = MessageBox.Show("入库成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);                       
                        if (DialogResult.OK == Rt)
                          {
                                this.Close();
                          }
                        }
                    else
                    {
                        return;
                    }
                    return;
                }
                else
                {
                    lblinfoinstock.ForeColor = Color.Red;
                    lblinfoinstock.Text = "超过了最大可入库数量:" + (int.Parse(enableqty == "" ? "0" : enableqty) - alreadyqty).ToString();
                    return;
                }


             // }
            //if (NGqty >= 0)
            //{
            //    string states = "NG";
            //    int stockqty = 0;

            //    if (txtstockqty.Text != "")
            //          stockqty = int.Parse(txtstockqty.Text);
            //    else
            //          stockqty = OKqty;
            //    string lotno = NGLotno;
            //    string remarks = NGremark;

            //    string org_id = dtems.Rows[0]["org_id"].ToString();
            //    string srecept = dtems.Rows[0]["deliveryid"].ToString();
            //    string productcode = dtems.Rows[0]["materialcode"].ToString();
            //    string item_id = dtems.Rows[0]["INVENTORY_ITEM_ID"].ToString();
            //    int alreadyqty = OtherReceiveInstock(org_id, srecept, productcode);
            //    string enableqty = dtems.Rows[0]["qty"].ToString();

            //    //还可入库数量(大于将要入库数量)
            //    if (int.Parse(enableqty == "" ? "0" : enableqty) - alreadyqty >= stockqty)
            //    {
            //            // string barcodemsg = IQC_TestOtherInstock(org_id, item_id, productcode, srecept, productcode, stockqty.ToString(), enableqty, Login.username, txtlotno.Text, states, remarks);

            //            string barcodemsg = insertBarCode(org_id, item_id, productcode, stockqty, Int32.Parse(enableqty), Login.username, srecept, productcode, lotno, states, remarks, conn, tran1);

            //        if (barcodemsg.IndexOf("成功") >= 0)
            //        {

            //            bool flag = DbAccess.ExecuteSql(sql);
            //            if (flag == false)
            //            {
            //                lblinfoinstock.ForeColor = Color.Red;
            //                lblinfoinstock.Text = "入库失败";
            //                tran1.Rollback();
            //                return;
            //            }
            //            tran1.Commit();
            //            lblinfoinstock.ForeColor = Color.Blue;
            //            lblinfoinstock.Text = barcodemsg;
            //            DialogResult Rt = MessageBox.Show("入库成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //                if (DialogResult.OK == Rt)
            //                {
            //                    this.Close();
            //                }

            //            }
            //        else
            //        {
            //            return;
            //        }
            //        return;
            //    }
            //    else
            //    {
            //        lblinfoinstock.ForeColor = Color.Red;
            //        lblinfoinstock.Text = "超过了最大可入库数量:" + (int.Parse(enableqty == "" ? "0" : enableqty) - alreadyqty).ToString();
            //        return;
            //    }

            //}
            }
            catch (Exception ex)
            {
                lblinfoinstock.Text = ex.ToString();
                tran1.Rollback();
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

        bool checkcuscode(string cuscode, string materialcode)
         {
             string checkcus = "";  
             string checkcuscodesql = @" select cuscode from OEM_EMSHYTCusRelation where hytcode='" + materialcode + "' order by  operdate desc  ";
             DataTable checkcuscodedt = Common.DbAccess.SelectBySql(checkcuscodesql).Tables[0];
            if (checkcuscodedt == null || checkcuscodedt.Rows.Count < 1)
            {
                return false;
            }
            else
            {
                checkcus = checkcuscodedt.Rows[0]["cuscode"].ToString().ToUpper();
                if (checkcus != cuscode)
                    return false;
                else
                    return true;
            }
         }


        private void txtproductname_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtproductname.Text != "")
            {      
                int k = 0;
                for (int i = 0; i < gridView.RowCount; i++)
                {
                    if (checkcuscode(txtproductname.Text.ToUpper(), gridView.GetRowCellValue(i, gridView.Columns["materialcode"]).ToString()))
                    {
                        labmessage.Text = "物料编码正确！";
                        labmessage.ForeColor = Color.Blue;
                        countcuscode = countcuscode + 1;
                        txtproductname.Text = "";
                        txtproductname.Focus();
                        break;
                    }
                    else
                        k++;
                }
                if (k == gridView.RowCount)
                {
                    labmessage.Text = "客户物料编码错误！";
                    labmessage.ForeColor = Color.Red;
                    countcuscode = 0; 
                    MessageBox.Show(txtproductname.Text + ",客户物料编码错误,请认真核对!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtproductname.Text = "";
                    txtproductname.Focus();
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