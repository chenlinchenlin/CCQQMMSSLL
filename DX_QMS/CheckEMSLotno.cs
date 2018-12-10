using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
//using System.Data.OracleClient;
//using Oracle.DataAccess.Client;
using DevExpress.XtraBars;
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.Data.OleDb;
using DevExpress.XtraEditors.Controls;
using Oracle.ManagedDataAccess.Client;

namespace DX_QMS
{
    public partial class CheckEMSLotno : DevExpress.XtraEditors.XtraForm
    {
        int StockQty = 0;
        string lotno="",  remarks = "", states="OK";
        string materialcode = "";
        string sMaterialname = "";
        public bool flag = false;
        public DataTable dtpub = null, dtalready = new DataTable();


        public CheckEMSLotno()
        {
            InitializeComponent();
        }

        public CheckEMSLotno(int stockqty,string Lotno, string txtremarks,string States,string Materialname, DataTable DtPub,DataTable DtCheck)
        {
            InitializeComponent();

            dtalready.Columns.Add("id", typeof(string));
            dtalready.Columns.Add("qty", typeof(int));

            StockQty = stockqty;   ////入库数量
            lotno = Lotno;        /////批次号
            remarks = txtremarks;  //////备注信息
            states = States;       //////入库结果
            dtpub = DtPub;         //////数据集
            sMaterialname = Materialname;
            gridControl.DataSource = DtCheck;
        }


        private DataTable IQCCheckResultLotno(string lotno)
        {
            DataTable dsEMSOther = null;

            string sSql = "select org_id,deliveryid,d.materialcode,materialname,qty,vendorname,lot_number from delivery d left join MaterialSpec m on d.materialcode=m.materialcode where lotno='" + lotno + "'";

            DataSet ds = DbAccess.SelectBySql(sSql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                Int64 x = 0;
                if (Int64.TryParse(ds.Tables[0].Rows[0]["deliveryid"].ToString(), out x))
                {
                    string sDeliveryID = ds.Tables[0].Rows[0]["deliveryid"].ToString();
                    string sMaterialCode = ds.Tables[0].Rows[0]["materialcode"].ToString();

                    string EMSSql = " select deliveryid, max(item) transactionid, round(sum(qty),0) qty,m.materialcode, max(m.materialname) materialname,";
                    EMSSql += " max(unit) unit, max(vendorname) vendorname, max(vendorcode) vendorcode,max(INVENTORY_ITEM_ID) INVENTORY_ITEM_ID,max(org_id) org_id,'' lot_number ";
                    EMSSql += " from deliveryEMSOtherRec d left join OEM_EMSHYTCusRelation o on d.vendorcode=o.cuscode left join  materialspec m on o.hytcode=m.materialcode  where  deliveryid='" + sDeliveryID + "' and m.materialcode='" + sMaterialCode + "'";
                    EMSSql += " group by deliveryid,m.materialcode";
                    dsEMSOther = DbAccess.SelectBySql(EMSSql).Tables[0];
                }
            }

            return dsEMSOther;

            //if (dsEMSOther != null && dsEMSOther.Rows.Count > 0)
            //{
            //    dtEMS = dsEMSOther.Tables[0];
            //    m = int.Parse(dsEMSOther.Tables[0].Rows[0]["qty"].ToString());
            //    txtstockqty.Text = dsEMSOther.Tables[0].Rows[0]["qty"].ToString();
            //    txttotalqty.Text = m.ToString();
            //}
        }


        private DataTable IQCCheckLotnoResult(string lotno)
        {
            string sql = " declare @recept varchar(30),@p varchar(30)";
            sql += "select @recept=receptid,@p=Productcode from IQC_TestList where LotNo='" + lotno + "'";
            sql += " SELECT  deliveryid,item, materialcode,d.lotno,(d.qty-isnull(finOK.qty,0)-isnull(finNG.qty,0)) qty,finOK.qty OKqty,finNG.qty NGqty,org_id,iqcTransactionid,'' TestResult,(d.qty-isnull(finOK.qty,0)-isnull(finNG.qty,0)) checkqty ,'' remark from delivery d left join";
            sql += "(select distinct lotno,productcode,receptid,states,isnull(sum(qty),0) qty from deliveryEMSCheckLotno t where receptid=@recept and productcode=@p and states='OK' group by lotno,productcode,receptid,states";
            sql += ") finOK on d.lotno=finOK.lotno and d.materialcode=finOK.productcode and d.deliveryid=finOK.receptid  left join ";
            sql += "(select lotno,productcode,receptid,states,isnull(sum(qty),0) qty from deliveryEMSCheckLotno t where receptid=@recept and productcode=@p and states='NG' group by lotno,productcode,receptid,states";
            sql += ") finNG on d.lotno=finNG.lotno and d.materialcode=finNG.productcode and d.deliveryid=finNG.receptid  ";
            sql += " where deliveryid=@recept and materialcode=@p ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }


        private void CheckEMSLotno_Load(object sender, EventArgs e)
        {
            //////////////////////////IQCCheckResultLotno("");

            ///// gridControl.DataSource = IQCCheckLotnoResult("Z000196691");

            //cbok.Enabled = false;
            //cbNG.Enabled = false;

            int CountQty = 0 , sstockqty = 0;   ///// StockQty
            sstockqty = StockQty;

            for (int i = 0; i < gridView.RowCount; i++)
            {
                if ( int.Parse(gridView.GetRowCellValue(i, gridView.Columns["checkqty"]).ToString()) > 0 )
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
            }




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

        void checklot(string lotno ,string materialcode)
        {

            try
            {
                string checkcodesql = @" select materialcode from delivery where lotno='"+ lotno + "' ";
                DataTable checkcodedt = Common.DbAccess.SelectBySql(checkcodesql).Tables[0];
                if (checkcodedt == null || checkcodedt.Rows.Count < 1)
                {
                    MessageBox.Show("未查询到我公司的编码，无法核对贴标正确性");
                     return;
                }
                string cuscode = checkcodedt.Rows[0]["materialcode"].ToString();
                string checkcuscodesql = @" select cuscode from OEM_EMSHYTCusRelation where hytcode='"+ cuscode + "' ";
                DataTable checkcuscodedt = Common.DbAccess.SelectBySql(checkcuscodesql).Tables[0];
                if (checkcuscodedt == null || checkcuscodedt.Rows.Count < 1)
                {
                    MessageBox.Show("未查询到相应的客户料号，无法核对贴标正确性");
                    return;
                }
                string materialcuscode = checkcuscodedt.Rows[0]["cuscode"].ToString();

                if (materialcuscode != materialcode)
                {
                    MessageBox.Show("所贴条码与料盘信息不对应，请重贴！");
                    return;
                }
            }
                catch
                {

                }
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
        /*
        public string insertOracle(Int32 tran_id, Int32 pqty, string baruser, string p_comments, OracleConnection cn, OracleTransaction tran)
        {
            OracleCommand sqlCmd;
            OracleParameter sqlParm;

            sqlCmd = new OracleCommand();
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.CommandText = "CUX_RCV_TRS_PKG.rcv_check";
            sqlCmd.Connection = cn;
            sqlCmd.Transaction = tran;


            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_rcv_transaction_id";
            sqlParm.DbType = DbType.Int32;
            sqlParm.Value = tran_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_qty";
            sqlParm.DbType = DbType.Int32;
            sqlParm.Value = pqty;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_barcode_user";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = baruser;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_comments";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = p_comments;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter("X_ERR_MSG", OracleType.VarChar, 50);
            sqlParm.Direction = ParameterDirection.Output;
            sqlCmd.Parameters.Add(sqlParm);

            sqlCmd.ExecuteNonQuery();
            return sqlCmd.Parameters["X_ERR_MSG"].Value.ToString();

        }
        */

        public string insertOracle(Int32 tran_id, Int32 pqty, string baruser, string p_comments, Oracle.ManagedDataAccess.Client.OracleConnection cn, Oracle.ManagedDataAccess.Client.OracleTransaction tran)
        {

            Oracle.ManagedDataAccess.Client.OracleCommand sqlCmd;
            Oracle.ManagedDataAccess.Client.OracleParameter sqlParm;

            sqlCmd = new Oracle.ManagedDataAccess.Client.OracleCommand();
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.CommandText = "CUX_RCV_TRS_PKG.RCV_CHECK";
            sqlCmd.Connection = cn;
            sqlCmd.Transaction = tran;

            //sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter("X_ERR_MSG", Oracle.ManagedDataAccess.Client.OracleDbType.NVarchar2, 50);
            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter("X_ERR_MSG", Oracle.ManagedDataAccess.Client.OracleDbType.NVarchar2, ParameterDirection.Output);
            //sqlParm.Direction = ParameterDirection.Output;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_rcv_transaction_id";
            //sqlParm.DbType = DbType.Int32;
            //sqlParm.DbType = DbType.VarNumeric;
            sqlParm.OracleDbType = OracleDbType.Int32;
            sqlParm.Value = tran_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_qty";
            //sqlParm.DbType = DbType.Int32;
            //sqlParm.DbType = DbType.VarNumeric;
            sqlParm.OracleDbType = OracleDbType.Int32;
            sqlParm.Value = pqty;
            sqlCmd.Parameters.Add(sqlParm);


            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_barcode_user";
            //sqlParm.DbType = DbType.String;
            //sqlParm.DbType = DbType.AnsiString;
            sqlParm.OracleDbType = OracleDbType.NVarchar2;
            sqlParm.Value = baruser;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_comments";
            //sqlParm.DbType = DbType.String;
            //sqlParm.DbType = DbType.AnsiString;
            sqlParm.OracleDbType = OracleDbType.NVarchar2;
            sqlParm.Value = p_comments;
            sqlCmd.Parameters.Add(sqlParm);

            ////sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter("X_ERR_MSG", Oracle.ManagedDataAccess.Client.OracleDbType.NVarchar2, 50);
            //sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter("X_ERR_MSG", Oracle.ManagedDataAccess.Client.OracleDbType.NVarchar2, ParameterDirection.Output);
            ////sqlParm.Direction = ParameterDirection.Output;
            //sqlCmd.Parameters.Add(sqlParm);

            sqlCmd.ExecuteNonQuery();
            return sqlCmd.Parameters["X_ERR_MSG"].Value.ToString();

        }

        /*
        public string insertOracleReject(Int32 tran_id, Int32 pqty, string p_inspection_code, Int32 p_reason_id, string p_comments, string baruser, OracleConnection cn, OracleTransaction tran)
        {

            OracleCommand sqlCmd;
            OracleParameter sqlParm;

            sqlCmd = new OracleCommand();
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.CommandText = "CUX_RCV_TRS_PKG.rcv_reject";
            sqlCmd.Connection = cn;
            sqlCmd.Transaction = tran;


            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_rcv_transaction_id";
            sqlParm.DbType = DbType.Int32;
            sqlParm.Value = tran_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_qty";
            sqlParm.DbType = DbType.Int32;
            sqlParm.Value = pqty;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_inspection_code";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = p_inspection_code;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_reason_id";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = p_reason_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_comments";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = p_comments;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_barcode_user";
            sqlParm.DbType = DbType.String;
            sqlParm.Value = baruser;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new OracleParameter("x_err_msg", OracleType.VarChar, 50);
            sqlParm.Direction = ParameterDirection.Output;
            sqlCmd.Parameters.Add(sqlParm);

            sqlCmd.ExecuteNonQuery();
            return sqlCmd.Parameters["x_err_msg"].Value.ToString();

        }
        */

        public string insertOracleReject(Int32 tran_id, Int32 pqty, string p_inspection_code, Int32 p_reason_id, string p_comments, string baruser, Oracle.ManagedDataAccess.Client.OracleConnection cn, Oracle.ManagedDataAccess.Client.OracleTransaction tran)
        {

            Oracle.ManagedDataAccess.Client.OracleCommand sqlCmd;
            Oracle.ManagedDataAccess.Client.OracleParameter sqlParm;

            sqlCmd = new Oracle.ManagedDataAccess.Client.OracleCommand();
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.CommandText = "CUX_RCV_TRS_PKG.RCV_REJECT";
            sqlCmd.Connection = cn;
            sqlCmd.Transaction = tran;

            //sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter("x_err_msg", Oracle.ManagedDataAccess.Client.OracleDbType.Varchar2, 50);
            //sqlParm.Direction = ParameterDirection.Output;
            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter("X_ERR_MSG", Oracle.ManagedDataAccess.Client.OracleDbType.NVarchar2, ParameterDirection.Output);
            sqlCmd.Parameters.Add(sqlParm);


            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_rcv_transaction_id";
            //sqlParm.DbType = DbType.Int32;
            sqlParm.OracleDbType = OracleDbType.Int32;
            sqlParm.Value = tran_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_qty";
            //sqlParm.DbType = DbType.Int32;
            sqlParm.OracleDbType = OracleDbType.Int32;
            sqlParm.Value = pqty;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_inspection_code";
            //sqlParm.DbType = DbType.String;
            sqlParm.OracleDbType = OracleDbType.NVarchar2;
            sqlParm.Value = p_inspection_code;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_reason_id";
            //sqlParm.DbType = DbType.String;
            sqlParm.OracleDbType = OracleDbType.Int32;
            sqlParm.Value = p_reason_id;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_comments";
            //sqlParm.DbType = DbType.String;
            sqlParm.OracleDbType = OracleDbType.NVarchar2;
            sqlParm.Value = p_comments;
            sqlCmd.Parameters.Add(sqlParm);

            sqlParm = new Oracle.ManagedDataAccess.Client.OracleParameter();
            sqlParm.Direction = ParameterDirection.Input;
            sqlParm.ParameterName = "p_barcode_user";
            //sqlParm.DbType = DbType.String;
            sqlParm.OracleDbType = OracleDbType.NVarchar2;
            sqlParm.Value = baruser;
            sqlCmd.Parameters.Add(sqlParm);
           
            sqlCmd.ExecuteNonQuery();
            return sqlCmd.Parameters["X_ERR_MSG"].Value.ToString(); 

        }


        private string getdeliveryinfo(string receipt, string materialcode)
        {
            int receiveqty = 0, acceptqty = 0;
            string sOraSql;
            //(状态为:RECEIVE,ACCEPT)
            sOraSql = " select transaction_type,ship_to_org_id,INVENTORY_ITEM_ID,transaction_id ,receipt_num, round(sum(quantity),0) primary_quantity, item_number,max(item_desc) item_desc,max(transaction_date) transaction_date,ORG_ID ";
            sOraSql += " from cux_inv_rec4_v where receipt_num='" + receipt + "' and item_number='" + materialcode + "'";
            sOraSql += " group by ship_to_org_id,transaction_id,INVENTORY_ITEM_ID,receipt_num,item_number,transaction_type,ORG_ID";

            DataSet dsERPDelivery = Common.DbAccess.SelectByOracle(sOraSql);
            if (dsERPDelivery != null && dsERPDelivery.Tables.Count > 0 && dsERPDelivery.Tables[0].Rows.Count > 0)
            {
                DataRow[] drr;
                //if(dsERPDelivery.Tables[0].Rows[0]["ORG_ID"].ToString()!="82")
                if (dsERPDelivery.Tables[0].Rows[0]["ORG_ID"].ToString() == "81")
                    drr = dsERPDelivery.Tables[0].Select("transaction_type='RECEIVE'");
                else if (dsERPDelivery.Tables[0].Rows[0]["ORG_ID"].ToString() == "203")
                    drr = dsERPDelivery.Tables[0].Select("transaction_type='TRANSFER'");
                else
                    drr = dsERPDelivery.Tables[0].Select("transaction_type='TRANSFER'");
                if (drr.Length > 0)
                {
                    for (int i = 0; i < drr.Length; i++)
                    {
                        receiveqty = receiveqty + int.Parse(drr[i]["primary_quantity"].ToString());
                    }
                }
                DataRow[] dra = dsERPDelivery.Tables[0].Select("transaction_type='ACCEPT'");
                if (dra.Length > 0)
                {
                    for (int j = 0; j < dra.Length; j++)
                    {
                        acceptqty = acceptqty + int.Parse(dra[j]["primary_quantity"].ToString());
                    }
                }
            }
            return "剩余量:" + receiveqty.ToString() + ",已检量：" + acceptqty.ToString();
        }

        bool updateMould(string productcode, string thisTimeQty)
        {
            string sql = "  if  exists(select 1 from IQC_MouldCycle where materialcode = '" + productcode + "' )   ";
            sql += "  update IQC_MouldCycle set RestQty = RestQty - " + thisTimeQty + " ,thisTimeQty = " + thisTimeQty + " ,updateMan = '" + Login.username + "' ,updateTime = GETDATE() ";
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

            if (gridView.GetSelectedRows().Length < 1)
                return;

            int OKqty = 0, NGqty = 0;
            string sql = "";
            string OKLotno = "", NGLotno = "";
            string OKremark = "", NGremark = "";

            for (int k = gridView.GetSelectedRows().Length; k > 0; k--)
            {
                DataRow db = gridView.GetDataRow(gridView.GetSelectedRows()[k - 1]);

                if (db["TestResult"].ToString() == "")
                {
                    continue;
                }

                if (int.Parse (db["qty"].ToString()) < int .Parse (db["checkqty"].ToString()))
                {
                    MessageBox.Show("确认数量不能大于可用数量");
                    return;
                }

                if (db["TestResult"].ToString() == "OK" && int.Parse(db["qty"].ToString())> 0)
                {
                    OKqty = OKqty + int.Parse(db["checkqty"].ToString());
                    OKLotno = db["lotno"].ToString();
                    OKremark = db["remark"].ToString();

                }
                if (db["TestResult"].ToString() == "NG" && int.Parse(db["qty"].ToString()) > 0)
                {
                    NGqty = NGqty + int.Parse(db["checkqty"].ToString());
                    NGLotno = db["lotno"].ToString();
                    NGremark = db["remark"].ToString();
                }

                sql += " if not exists(select 1 from deliveryEMSCheckLotno where lotno='" + db["lotno"].ToString() + "' and states='" + db["TestResult"].ToString() + "')";
                sql += " insert into deliveryEMSCheckLotno(org_id, receptid,productcode,qty, lotno, states,eventuser, eventtime)";
                sql += "values('" + db["org_id"].ToString() + "','" + db["deliveryid"].ToString() + "','" + db["materialcode"].ToString() + "','" + db["checkqty"].ToString() + "','" + db["lotno"].ToString() + "','" + db["TestResult"].ToString() + "','" + Login.userId + "',getdate())";
                sql += " else update deliveryEMSCheckLotno set qty=qty+" + db["checkqty"].ToString() + " where lotno='" + db["lotno"].ToString() + "' and states='" + db["TestResult"].ToString() + "'";

                if (!sMaterialname.Contains("【F1"))
                {
                    string connPANASql = "server=10.100.64.4,6815;database=PanaCIM;user id=sa;password=PANASONIC1!;";
                    SqlConnection sqlconn = new SqlConnection(connPANASql);
                    if (sqlconn.State == ConnectionState.Closed)
                        sqlconn.Open();
                    SqlTransaction tran = sqlconn.BeginTransaction();
                    DataTable Uploaddt = GetLotList(db["lotno"].ToString());
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
                            //this.lblMessage.Text = "保存成功!" + ",松下系统,已经存在该数据!" + Uploaddt.Rows[0]["reelid"].ToString();
                            tran.Rollback();
                        }
                        else
                        {
                           // this.lblMessage.Text = Uploaddt.Rows[0]["reelid"].ToString() + "保存成功!" + ",写入松下数据失败!" + result + ",请重试一次";
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

            if (OKqty > 0)
            {
                states = "OK";
                int stockqty = OKqty;
                lotno = OKLotno;

                if (OKremark != "" )
                  remarks = OKremark;


               DataTable dt = dtpub;

                SqlConnection conn = new SqlConnection(DbAccess.connSql);
                //OracleConnection orac = new OracleConnection(DbAccess.connOral);
                Oracle.ManagedDataAccess.Client.OracleConnection orac = new Oracle.ManagedDataAccess.Client.OracleConnection(DbAccess.connOral);
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
                if (orac.State == ConnectionState.Closed)
                    orac.Open();
                SqlTransaction tran1 = conn.BeginTransaction();
                // OracleTransaction oratran = orac.BeginTransaction();
                Oracle.ManagedDataAccess.Client.OracleTransaction oratran = orac.BeginTransaction();
                try
                {

                    Int32 org_id = 0, item_id = 0, tran_id = 0;
                    Int64 recepid = 0;
                    string org_idnew = "";
                    int instockqty = 0;
                    ////////// int restenableqty = int.Parse(txtstockqty.Text);

                    int restenableqty = stockqty;


                    string msg = "", barcodemsg = "", msgcurrent = "", barcodemsgcurrent = "";
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        int alreadyqty = 0;
                        org_idnew = dt.Rows[j]["ORGANIZATION_CODE"].ToString();
                        org_id = Int32.Parse(dt.Rows[j]["ship_to_org_id"].ToString());
                        item_id = Int32.Parse(dt.Rows[j]["INVENTORY_ITEM_ID"].ToString());
                        tran_id = Int32.Parse(dt.Rows[j]["transaction_id"].ToString());
                        recepid = Int64.Parse(dt.Rows[0]["receipt_num"].ToString());

                        //判断当前行可检验数量大于界面录入数量,则把所有数量录入到第一笔.
                        if (int.Parse(dt.Rows[j]["primary_quantity"].ToString()) >= restenableqty)
                        {
                            //则把用户输入的检验数量作为此行的检验数
                            instockqty = restenableqty;
                            //20150312新增
                            //x表示还在接口表中的数量
                            //int x = ERPinterface(tran_id.ToString());
                            //用临时变量记录此次操作数量
                            int x = alreadyqty;
                            //y表示还可处理的数量
                            int y = int.Parse(dt.Rows[j]["primary_quantity"].ToString()) - x;
                            if (y >= restenableqty)
                                instockqty = restenableqty;
                            else if (y > 0 && y < restenableqty)
                                instockqty = y;
                            else
                                continue;
                            //20150312新增

                            barcodemsg = barcodemsg + insertBarCode(org_idnew.ToString(), item_id.ToString(), tran_id.ToString(), instockqty, Int32.Parse(dt.Rows[j]["primary_quantity"].ToString()),
                                                                     Login.username, recepid.ToString(), dt.Rows[j]["item_number"].ToString(), lotno, states, remarks, conn, tran1);
                            if (barcodemsg.IndexOf("成功") >= 0)
                            {
                                if (states == "OK")
                                {
                                    msgcurrent = insertOracle(tran_id, instockqty, Login.username, remarks, orac, oratran);
                                    try
                                    {
                                        updateMould(dt.Rows[j]["item_number"].ToString(), instockqty.ToString());
                                    }
                                    catch
                                    { }
                                    if (msgcurrent.IndexOf("Error") >= 0)
                                    {
                                        msg = msg + msgcurrent;

                                        break;
                                    }
                                    else
                                    {
                                        msg = msg + msgcurrent;
                                        //////////// txtstockqty.Text = "0";
                                        stockqty = 0;

                                        break;
                                    }
                                }
                                else if (states == "NG")
                                {
                                    msgcurrent = insertOracleReject(tran_id, instockqty, "BAD", 31, remarks, Login.username, orac, oratran);
                                    if (msgcurrent.IndexOf("Error") >= 0)
                                    {
                                        msg = msg + msgcurrent;
                                        break;
                                    }
                                    else
                                    {
                                        msg = msg + msgcurrent;
                                        ///////////// txtstockqty.Text = "0";
                                        stockqty = 0;

                                        break;
                                    }
                                }
                            }
                            else
                            {
                                //防止第一笔已入库(一批次分多次入库问题),要继续循环下一笔
                                continue;
                            }
                        }
                        //如果当前行的数量小于录入数量,则把当前行先进行检验通过该行的数量,同时把总共要检验的数量-第一笔已录的=作为剩下的可录入数量，再进行下一次循环
                        else
                        {
                            instockqty = int.Parse(dt.Rows[j]["primary_quantity"].ToString());
                            //20150312新增
                            //x表示还在接口表中的数量
                            //int x = ERPinterface(tran_id.ToString());
                            //用临时变量记录此次操作数量

                            DataRow[] rwalready = dtalready.Select(" id='" + tran_id + "'");
                            if (rwalready.Length > 0)
                            {
                                for (int z = 0; z < rwalready.Length; z++)
                                {
                                    alreadyqty = alreadyqty + int.Parse(rwalready[z]["qty"].ToString());
                                }
                            }
                            int x = alreadyqty;
                            //y表示还可处理的数量
                            int y = int.Parse(dt.Rows[j]["primary_quantity"].ToString()) - x;
                            if (y >= restenableqty)
                                instockqty = restenableqty;
                            else if (y > 0 && y < restenableqty)
                                instockqty = y;
                            else
                                continue;
                            //20150312新增

                            barcodemsgcurrent = insertBarCode(org_idnew.ToString(), item_id.ToString(), tran_id.ToString(), instockqty, instockqty,
                                                                    Login.username, recepid.ToString(), dt.Rows[j]["item_number"].ToString(), lotno, states, remarks, conn, tran1);
                            barcodemsg = barcodemsg + barcodemsgcurrent;
                            if (barcodemsgcurrent.IndexOf("成功") >= 0)
                            {
                                if (states == "OK")
                                {
                                    msgcurrent = insertOracle(tran_id, instockqty, Login.username, remarks, orac, oratran);
                                    try
                                    {
                                        updateMould(dt.Rows[j]["item_number"].ToString(), instockqty.ToString());
                                    }
                                    catch
                                    { }
                                    if (msgcurrent.IndexOf("Error") >= 0)
                                    {
                                        msg = msg + msgcurrent;
                                        break;
                                    }
                                    msg = msg + msgcurrent;
                                    restenableqty = restenableqty - instockqty;
                                    ////////////  txtstockqty.Text = restenableqty.ToString();
                                    stockqty = restenableqty;

                                    DataRow newRow;
                                    newRow = dtalready.NewRow();
                                    newRow["id"] = tran_id;
                                    newRow["qty"] = instockqty.ToString();
                                    dtalready.Rows.Add(newRow);

                                    continue;
                                }
                                else if (states == "NG")
                                {
                                    msgcurrent = insertOracleReject(tran_id, instockqty, "BAD", 31, remarks, Login.username, orac, oratran);
                                    if (msgcurrent.IndexOf("Error") >= 0)
                                    {
                                        msg = msg + msgcurrent;
                                        break;
                                    }
                                    msg = msg + msgcurrent;
                                    restenableqty = restenableqty - instockqty;
                                    ///////////  txtstockqty.Text = restenableqty.ToString();
                                    stockqty = restenableqty;

                                    DataRow newRow;
                                    newRow = dtalready.NewRow();
                                    newRow["id"] = tran_id;
                                    newRow["qty"] = instockqty.ToString();
                                    dtalready.Rows.Add(newRow);
                                    continue;
                                }
                            }
                            else
                            {
                                //防止第一笔已入库(一批次分多次入库问题),要继续循环下一笔
                                continue;
                            }
                        }
                    }

                    if (msg.IndexOf("Error") >= 0)
                    {
                        this.lblinfoinstock.ForeColor = Color.Red;
                        lblinfoinstock.Text = msg + ";ERP写入失败!";
                        tran1.Rollback();
                        oratran.Rollback();
                    }
                    else if (barcodemsg.IndexOf("失败") >= 0)
                    {
                        lblinfoinstock.ForeColor = Color.Red;
                        lblinfoinstock.Text = barcodemsg + ";ERP写入失败!";
                        tran1.Rollback();
                        oratran.Rollback();
                    }
                    else
                    {

                        bool flag = DbAccess.ExecuteSql(sql);
                        if (flag == false)
                        {
                            lblinfoinstock.ForeColor = Color.Red;
                            lblinfoinstock.Text = barcodemsg + ";ERP写入失败!";
                            tran1.Rollback();
                            oratran.Rollback();
                            return;
                        }
                        else
                        {
                          
                            lblinfoinstock.ForeColor = Color.Blue;
                            ////////txtstockqty.Text = "";
                            stockqty = 0;
                            ////////this.txtremarks.Text = "";
                            tran1.Commit();
                            oratran.Commit();
                            lblinfoinstock.Text = states + ";" + getdeliveryinfo(recepid.ToString(), dt.Rows[0]["item_number"].ToString());
                            DialogResult Rt = MessageBox.Show("入库OK", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (DialogResult.OK == Rt)
                            {
                                this.Close();
                            }                          
                            //dtpub = null;

                        }
                    
                    }
                }
                catch (Exception ex)
                {
                    lblinfoinstock.Text = ex.ToString();
                    tran1.Rollback();
                    oratran.Rollback();
                }

            }

            if (NGqty > 0)
            {
                states = "NG";
                int stockqty = NGqty;
                lotno = NGLotno;

                if ( NGremark != "" )
                   remarks = NGremark;


                DataTable dt = dtpub;

                SqlConnection conn = new SqlConnection(DbAccess.connSql);
                //OracleConnection orac = new OracleConnection(DbAccess.connOral);
                Oracle.ManagedDataAccess.Client.OracleConnection orac = new Oracle.ManagedDataAccess.Client.OracleConnection(DbAccess.connOral);
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
                if (orac.State == ConnectionState.Closed)
                    orac.Open();
                SqlTransaction tran1 = conn.BeginTransaction();
                //OracleTransaction oratran = orac.BeginTransaction();
                Oracle.ManagedDataAccess.Client.OracleTransaction oratran = orac.BeginTransaction();
                try
                {

                    Int32 org_id = 0, item_id = 0, tran_id = 0;
                    Int64 recepid = 0;
                    string org_idnew = "";
                    int instockqty = 0;
                    ////////// int restenableqty = int.Parse(txtstockqty.Text);

                    int restenableqty = stockqty;

                    string msg = "", barcodemsg = "", msgcurrent = "", barcodemsgcurrent = "";
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        int alreadyqty = 0;
                        org_idnew = dt.Rows[j]["ORGANIZATION_CODE"].ToString();
                        org_id = Int32.Parse(dt.Rows[j]["ship_to_org_id"].ToString());
                        item_id = Int32.Parse(dt.Rows[j]["INVENTORY_ITEM_ID"].ToString());
                        tran_id = Int32.Parse(dt.Rows[j]["transaction_id"].ToString());
                        recepid = Int64.Parse(dt.Rows[0]["receipt_num"].ToString());

                        //判断当前行可检验数量大于界面录入数量,则把所有数量录入到第一笔.
                        if (int.Parse(dt.Rows[j]["primary_quantity"].ToString()) >= restenableqty)
                        {
                            //则把用户输入的检验数量作为此行的检验数
                            instockqty = restenableqty;
                            //20150312新增
                            //x表示还在接口表中的数量
                            //int x = ERPinterface(tran_id.ToString());
                            //用临时变量记录此次操作数量
                            int x = alreadyqty;
                            //y表示还可处理的数量
                            int y = int.Parse(dt.Rows[j]["primary_quantity"].ToString()) - x;
                            if (y >= restenableqty)
                                instockqty = restenableqty;
                            else if (y > 0 && y < restenableqty)
                                instockqty = y;
                            else
                                continue;
                            //20150312新增

                            barcodemsg = barcodemsg + insertBarCode(org_idnew.ToString(), item_id.ToString(), tran_id.ToString(), instockqty, Int32.Parse(dt.Rows[j]["primary_quantity"].ToString()),
                                                                     Login.username, recepid.ToString(), dt.Rows[j]["item_number"].ToString(), lotno, states, remarks, conn, tran1);
                            if (barcodemsg.IndexOf("成功") >= 0)
                            {
                                if (states == "OK")
                                {
                                    msgcurrent = insertOracle(tran_id, instockqty, Login.username, remarks, orac, oratran);
                                    try
                                    {
                                        updateMould(dt.Rows[j]["item_number"].ToString(), instockqty.ToString());
                                    }
                                    catch {
                                    }
                                    if (msgcurrent.IndexOf("Error") >= 0)
                                    {
                                        msg = msg + msgcurrent;

                                        break;
                                    }
                                    else
                                    {
                                        msg = msg + msgcurrent;
                                      //////////  txtstockqty.Text = "0";
                                        break;
                                    }
                                }
                                else if (states == "NG")
                                {
                                    msgcurrent = insertOracleReject(tran_id, instockqty, "BAD", 31, remarks, Login.username, orac, oratran);
                                    if (msgcurrent.IndexOf("Error") >= 0)
                                    {
                                        msg = msg + msgcurrent;
                                        break;
                                    }
                                    else
                                    {
                                        msg = msg + msgcurrent;
                                      ///////////  txtstockqty.Text = "0";

                                        break;
                                    }
                                }
                            }
                            else
                            {
                                //防止第一笔已入库(一批次分多次入库问题),要继续循环下一笔
                                continue;
                            }
                        }
                        //如果当前行的数量小于录入数量,则把当前行先进行检验通过该行的数量,同时把总共要检验的数量-第一笔已录的=作为剩下的可录入数量，再进行下一次循环
                        else
                        {
                            instockqty = int.Parse(dt.Rows[j]["primary_quantity"].ToString());
                            //20150312新增
                            //x表示还在接口表中的数量
                            //int x = ERPinterface(tran_id.ToString());
                            //用临时变量记录此次操作数量

                            DataRow[] rwalready = dtalready.Select(" id='" + tran_id + "'");
                            if (rwalready.Length > 0)
                            {
                                for (int z = 0; z < rwalready.Length; z++)
                                {
                                    alreadyqty = alreadyqty + int.Parse(rwalready[z]["qty"].ToString());
                                }
                            }
                            int x = alreadyqty;
                            //y表示还可处理的数量
                            int y = int.Parse(dt.Rows[j]["primary_quantity"].ToString()) - x;
                            if (y >= restenableqty)
                                instockqty = restenableqty;
                            else if (y > 0 && y < restenableqty)
                                instockqty = y;
                            else
                                continue;
                            //20150312新增

                            barcodemsgcurrent = insertBarCode(org_idnew.ToString(), item_id.ToString(), tran_id.ToString(), instockqty, instockqty,
                                                                    Login.username, recepid.ToString(), dt.Rows[j]["item_number"].ToString(), lotno, states, remarks, conn, tran1);
                            barcodemsg = barcodemsg + barcodemsgcurrent;
                            if (barcodemsgcurrent.IndexOf("成功") >= 0)
                            {
                                if (states == "OK")
                                {
                                    msgcurrent = insertOracle(tran_id, instockqty, Login.username, remarks, orac, oratran);
                                    try
                                    {
                                        updateMould(dt.Rows[j]["item_number"].ToString(), instockqty.ToString());
                                    }
                                    catch {
                                    }
                                    if (msgcurrent.IndexOf("Error") >= 0)
                                    {
                                        msg = msg + msgcurrent;
                                        break;
                                    }
                                    msg = msg + msgcurrent;
                                    restenableqty = restenableqty - instockqty;
                                   ///////////// txtstockqty.Text = restenableqty.ToString();

                                    DataRow newRow;
                                    newRow = dtalready.NewRow();
                                    newRow["id"] = tran_id;
                                    newRow["qty"] = instockqty.ToString();
                                    dtalready.Rows.Add(newRow);

                                    continue;
                                }
                                else if (states == "NG")
                                {
                                    msgcurrent = insertOracleReject(tran_id, instockqty, "BAD", 31, remarks, Login.username, orac, oratran);
                                    if (msgcurrent.IndexOf("Error") >= 0)
                                    {
                                        msg = msg + msgcurrent;
                                        break;
                                    }
                                    msg = msg + msgcurrent;
                                    restenableqty = restenableqty - instockqty;
                                  ////////////  txtstockqty.Text = restenableqty.ToString();

                                    DataRow newRow;
                                    newRow = dtalready.NewRow();
                                    newRow["id"] = tran_id;
                                    newRow["qty"] = instockqty.ToString();
                                    dtalready.Rows.Add(newRow);
                                    continue;
                                }
                            }
                            else
                            {
                                //防止第一笔已入库(一批次分多次入库问题),要继续循环下一笔
                                continue;
                            }
                        }
                    }

                    if (msg.IndexOf("Error") >= 0)
                    {
                        this.lblinfoinstock.ForeColor = Color.Red;
                        lblinfoinstock.Text = msg + ";ERP写入失败!";
                        tran1.Rollback();
                        oratran.Rollback();
                    }
                    else if (barcodemsg.IndexOf("失败") >= 0)
                    {
                        lblinfoinstock.ForeColor = Color.Red;
                        lblinfoinstock.Text = barcodemsg + ";ERP写入失败!";
                        tran1.Rollback();
                        oratran.Rollback();
                    }
                    else
                    {
                        bool flag = DbAccess.ExecuteSql(sql);
                        if (flag == false)
                        {
                            lblinfoinstock.ForeColor = Color.Red;
                            lblinfoinstock.Text = barcodemsg + ";ERP写入失败!";
                            tran1.Rollback();
                            oratran.Rollback();
                            return;
                        }
                        else
                        {

                            lblinfoinstock.ForeColor = Color.Blue;
                            ////////txtstockqty.Text = "";
                            stockqty = 0;
                            ////////this.txtremarks.Text = "";
                            tran1.Commit();
                            oratran.Commit();
                            lblinfoinstock.Text = states + ";" + getdeliveryinfo(recepid.ToString(), dt.Rows[0]["item_number"].ToString());

                            DialogResult Rt = MessageBox.Show("入库NG", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (DialogResult.OK == Rt)
                            {
                                this.Close();
                            }
                            //dtpub = null;

                        }
                    }
                }
                catch (Exception ex)
                {
                    lblinfoinstock.Text = ex.ToString();
                    tran1.Rollback();
                    oratran.Rollback();
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

        private void txtproductcode_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && txtproductcode.Text != "")
            {
                int k = 0;
                for (int i = 0; i < gridView.RowCount; i++)
                {
                    if (txtproductcode.Text.Trim().ToUpper() == gridView.GetRowCellValue(i, gridView.Columns["materialcode"]).ToString())
                    {
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

        private void txtloto_KeyUp(object sender, KeyEventArgs e)
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


    }
}