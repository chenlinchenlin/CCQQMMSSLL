using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using System.Data;
using System.Data.SqlClient;
//using System.Data.OracleClient;
using Oracle.ManagedDataAccess.Client;
using DX_QMS.Common;

namespace DX_QMS
{
    public partial class TestIQCCheckAgain : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private string tranid = "", item_id = "", org_id = "";
        private string srecept = "", productcode = "";
        private string checktype = "NGToOK";


        public TestIQCCheckAgain()
        {
            InitializeComponent();
        }



        private void txtchecktype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtchecktype.Text == "由NG改OK")
                checktype = "NGToOK";
            else if (txtchecktype.Text == "由OK改NG")
                checktype = "OKToNG";
            else if (txtchecktype.Text == "由NG改特采")
                checktype = "NGTo特采";
            this.txtchecktype.Enabled = false;
            this.txtlotno.Focus();
        }

        private DataTable IfWriteToERP(string lotno)
        {
            int m = 0;
            DataTable dtpub = null;
            string sSql = "select deliveryid,materialcode,qty,vendorname,lot_number from delivery  where lotno='" + lotno + "'";
            DataSet ds = DbAccess.SelectBySql(sSql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                Int64 x = 0;
                //客供料下面的视图中没有
                if (Int64.TryParse(ds.Tables[0].Rows[0]["deliveryid"].ToString(), out x))
                {
                    string sDeliveryID = ds.Tables[0].Rows[0]["deliveryid"].ToString();
                    string sMaterialCode = ds.Tables[0].Rows[0]["materialcode"].ToString();
                    string sOraSql;
                   
                    sOraSql = " select ship_to_org_id,INVENTORY_ITEM_ID,transaction_id ,receipt_num, round(sum(quantity),0) primary_quantity, item_number,max(item_desc) item_desc,max(transaction_date) transaction_date,ORG_ID,transaction_type,ORGANIZATION_CODE,SHIP_TO_ORG_CODE,PARENT_TRANSACTION_TYPE ";
                    sOraSql += " from cux_inv_rec4_v where (transaction_type='ACCEPT' or transaction_type='REJECT') and receipt_num='" + sDeliveryID + "' and item_number='" + sMaterialCode + "'";
                    sOraSql += " group by ship_to_org_id,transaction_id,INVENTORY_ITEM_ID,receipt_num,item_number,ORG_ID,transaction_type,ORGANIZATION_CODE,SHIP_TO_ORG_CODE,PARENT_TRANSACTION_TYPE";

                    DataSet dsERPDelivery = DbAccess.SelectByOracle(sOraSql);
                    if (dsERPDelivery != null && dsERPDelivery.Tables.Count > 0 && dsERPDelivery.Tables[0].Rows.Count > 0)
                    {
                        if (checktype == "NGToOK")
                        {
                            DataRow[] rw = dsERPDelivery.Tables[0].Select("transaction_type='REJECT' ");
                            dtpub = dsERPDelivery.Tables[0].Clone();
                            for (int i = 0; i < rw.Length; i++)
                            {
                                dtpub.Rows.Add(rw[i].ItemArray);
                            }
                            for (int i = 0; i < dtpub.Rows.Count; i++)
                            {
                                m = m + int.Parse(dtpub.Rows[i]["primary_quantity"].ToString());
                            }
                        }
                        else if (checktype == "OKToNG")
                        {
                            DataRow[] rw = dsERPDelivery.Tables[0].Select("transaction_type='ACCEPT' ");
                            dtpub = dsERPDelivery.Tables[0].Clone();
                            for (int i = 0; i < rw.Length; i++)
                            {
                                dtpub.Rows.Add(rw[i].ItemArray);
                            }
                            for (int i = 0; i < dtpub.Rows.Count; i++)
                            {
                                m = m + int.Parse(dtpub.Rows[i]["primary_quantity"].ToString());
                            }
                        }
                    }
                    else
                    {
                        lblinfo.ForeColor = Color.Red;
                        lblinfo.Text = sDeliveryID + "还没有NG的记录,不能改判!";
                    }
                }
            }
            return dtpub;
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

        private void clearform()
        {
            txtchecktype.Enabled = true;
            txtlotno.Enabled = true;
            txtlotno.Text = "";
            txtlotno.Focus();
            txtenableqty.Text = "";
            txtqty.Text = "";
            txtreason.Text = "";
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


        private void IQCReCheck(DataTable dtpub)
        {
            int m = 0;
            string states = "";
            if (checktype == "NGToOK")
                states = "OK";
            else if (checktype == "OKToNG")
                states = "NG";
            else if (checktype == "NGTo特采")
                states = "OK";

            string remarks = "", noteid = "";
            if (checktype == "NGTo特采")
            {
                remarks = txtnoteid.Text + checktype + txtreason.Text;
                noteid = txtnoteid.Text;
            }
            else
                remarks = checktype + txtreason.Text;


            DataTable dt = dtpub;
            DataTable dtalready = new DataTable();

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
                int restenableqty = int.Parse(txtqty.Text);
                string msg = "", barcodemsg = "", msgcurrent = "", barcodemsgcurrent = "";
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    int alreadyqty = 0;
                    org_idnew = dt.Rows[j]["ORGANIZATION_CODE"].ToString();
                    org_id = Int32.Parse(dt.Rows[j]["ship_to_org_id"].ToString());
                    item_id = Int32.Parse(dt.Rows[j]["INVENTORY_ITEM_ID"].ToString());
                    tran_id = Int32.Parse(dt.Rows[j]["transaction_id"].ToString());
                    //recepid = Int32.Parse(dt.Rows[0]["receipt_num"].ToString());
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
                                                                 Login.username, recepid.ToString(), dt.Rows[j]["item_number"].ToString(), txtlotno.Text, "checkagain", remarks, conn, tran1);
                        if (barcodemsg.IndexOf("成功") >= 0)
                        {
                            if (states == "OK")
                            {
                                msgcurrent = insertOracle(tran_id, instockqty, Login.username, noteid + txtreason.Text, orac, oratran);
                                if (msgcurrent.IndexOf("Error") >= 0)
                                {
                                    msg = msg + msgcurrent;

                                    break;
                                }
                                else
                                {
                                    msg = msg + msgcurrent;
                                    txtqty.Text = "0";
                                    break;
                                }
                            }
                            else if (states == "NG")
                            {
                                msgcurrent = insertOracleReject(tran_id, instockqty, "BAD", 31, txtreason.Text, Login.username, orac, oratran);
                                if (msgcurrent.IndexOf("Error") >= 0)
                                {
                                    msg = msg + msgcurrent;
                                    break;
                                }
                                else
                                {
                                    msg = msg + msgcurrent;
                                    txtqty.Text = "0";
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

                        try
                        {
                            DataRow[] rwalready = dtalready.Select(" id='" + tran_id + "'");
                            if (rwalready.Length > 0)
                            {
                                for (int z = 0; z < rwalready.Length; z++)
                                {
                                    alreadyqty = alreadyqty + int.Parse(rwalready[z]["qty"].ToString());
                                }
                            }
                        }
                        catch
                        {
                            dtalready.Columns.Add("id", typeof(string));
                            dtalready.Columns.Add("qty", typeof(int));
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
                                                                Login.username, recepid.ToString(), dt.Rows[j]["item_number"].ToString(), txtlotno.Text, "checkagain", remarks, conn, tran1);
                        barcodemsg = barcodemsg + barcodemsgcurrent;
                        if (barcodemsgcurrent.IndexOf("成功") >= 0)
                        {
                            if (states == "OK")
                            {
                                msgcurrent = insertOracle(tran_id, instockqty, Login.username, noteid + txtreason.Text, orac, oratran);
                                if (msgcurrent.IndexOf("Error") >= 0)
                                {
                                    msg = msg + msgcurrent;
                                    break;
                                }
                                msg = msg + msgcurrent;
                                restenableqty = restenableqty - instockqty;
                                txtqty.Text = restenableqty.ToString();

                                DataRow newRow;
                                newRow = dtalready.NewRow();
                                newRow["id"] = tran_id;
                                newRow["qty"] = instockqty.ToString();
                                dtalready.Rows.Add(newRow);

                                continue;
                            }
                            else if (states == "NG")
                            {
                                msgcurrent = insertOracleReject(tran_id, instockqty, "BAD", 31, txtreason.Text, Login.username, orac, oratran);
                                if (msgcurrent.IndexOf("Error") >= 0)
                                {
                                    msg = msg + msgcurrent;
                                    break;
                                }
                                msg = msg + msgcurrent;
                                restenableqty = restenableqty - instockqty;
                                txtqty.Text = restenableqty.ToString();

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
                    this.lblinfo.ForeColor = Color.Red;
                    lblinfo.Text = msg + ";ERP写入失败!";
                    tran1.Rollback();
                    oratran.Rollback();
                }
                else if (barcodemsg.IndexOf("失败") >= 0)
                {
                    lblinfo.ForeColor = Color.Red;
                    lblinfo.Text = barcodemsg + ";ERP写入失败!";
                    tran1.Rollback();
                    oratran.Rollback();
                }
                else
                {
                    lblinfo.ForeColor = Color.Blue;
                    lblinfo.Text = "改判成功" + barcodemsg + ";" + msg;
                    txtqty.Text = "";
                    this.txtreason.Text = "";
                    tran1.Commit();
                    oratran.Commit();
                }


            }
            catch (Exception ex)
            {
                lblinfo.Text = ex.ToString();
                tran1.Rollback();
                oratran.Rollback();
            }

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (checktype == "NGTo特采" && txtnoteid.Text == "")
            {
                lblinfo.Text = "需要录入电子流号";
                lblinfo.ForeColor = Color.Red;
                DialogResult Rt = MessageBox.Show("电子流号为空，是否继续？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (DialogResult.Cancel == Rt)
                {
                    return;
                }
            }
            int i = 0;
            if (!int.TryParse(txtqty.Text.Trim(), out i))
                return;
            DataTable dt = IfWriteToERP(txtlotno.Text);
            if (dt == null || dt.Rows.Count <= 0)
            {
                if (int.Parse(txtqty.Text.Trim()) <= int.Parse(txtenableqty.Text))
                {
                    string remarks = "";
                    if (checktype == "NGTo特采")
                    {
                        remarks = txtnoteid.Text + checktype + txtreason.Text;
                    }
                    else
                        remarks = checktype + txtreason.Text;                   
                    string barcodemsg = IQC_TestOtherInstock(org_id, item_id, productcode, srecept, productcode, txtqty.Text, txtenableqty.Text, Login.username, txtlotno.Text.Trim(), "checkagain", remarks);
                    if (barcodemsg.IndexOf("成功") >= 0)
                    {
                        lblinfo.ForeColor = Color.Blue;
                        lblinfo.Text = barcodemsg;
                        clearform();
                    }
                    else
                    {
                        lblinfo.Text = barcodemsg;
                        lblinfo.ForeColor = Color.Red;
                    }
                }
                else
                {
                    lblinfo.ForeColor = Color.Red;
                    lblinfo.Text = "超过了最大可改判数量:" + this.txtenableqty.Text;
                    return;
                }
            }
            else
            {
                //对有接收单号的来料进行改判
                IQCReCheck(dt);
            }
        }

        private void btnreset_Click(object sender, EventArgs e)
        {
            clearform();
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

        private void txtlotno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                if (checktype == "NGToOK")
                {
                    string sqlorg = "select d.org_id,materialname from delivery d inner join MaterialSpec m on d.materialcode=m.materialcode where lotno='" + txtlotno.Text + "'";
                    DataTable dtorg = DbAccess.SelectBySql(sqlorg).Tables[0];
                    if (dtorg.Rows[0][0].ToString() == "HCL" && !dtorg.Rows[0]["materialname"].ToString().ToUpper().Contains("HBTG"))
                    {
                        DataTable dtToERP = IQCCheckResultToERP(txtlotno.Text);
                        if (dtToERP.Rows.Count > 0)
                        {
                            TestResultNGToOKCheck TRC = new TestResultNGToOKCheck(dtToERP);
                            TRC.ShowDialog();
                            return;
                        }
                        else
                        {
                            lblinfo.ForeColor = Color.Red;
                            lblinfo.Text = txtlotno.Text + "没有不良数或已改判完成!";
                            txtlotno.Focus();
                            txtlotno.Text = "";
                            txtlotno.Enabled = true;
                            return;
                        }
                    }

                    string sql = "select sum(qty) NGqty,max(transaction_id) transaction_id,max(ITEM_ID) ITEM_ID,max(org_id) org_id,receptid, productcode from Warehouse_IQCCheck c " ;
                    sql += " where exists(select receptid,productcode from Warehouse_IQCCheck where states='NG' and lotno='"+ txtlotno.Text.Trim()+ "' " ;
                    sql += "  and c.receptid=receptid and c.productcode=productcode) and states='NG' group by receptid,productcode ";
                    DataSet ds = DbAccess.SelectBySql(sql);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        tranid = ds.Tables[0].Rows[0]["NGqty"].ToString();
                        item_id = ds.Tables[0].Rows[0]["ITEM_ID"].ToString();
                        org_id = ds.Tables[0].Rows[0]["org_id"].ToString();
                        srecept = ds.Tables[0].Rows[0]["receptid"].ToString();
                        productcode = ds.Tables[0].Rows[0]["productcode"].ToString();

                        txtenableqty.Text = ds.Tables[0].Rows[0]["NGqty"].ToString();
                        txtqty.Focus();
                        txtlotno.Enabled = false;
                    }
                    else
                    {
                        lblinfo.ForeColor = Color.Red;
                        lblinfo.Text = txtlotno.Text + "不存在或没有不良数";
                        txtlotno.Focus();
                        txtlotno.Text = "";
                        txtlotno.Enabled = true;
                    }
                }
                else
                {
                    string sql = "select sum(qty) OKqty,max(transaction_id) transaction_id,max(ITEM_ID) ITEM_ID,max(org_id) org_id,receptid, productcode from Warehouse_IQCCheck c ";
                    sql += "  where exists(select receptid,productcode from Warehouse_IQCCheck where states='OK' and lotno='"+txtlotno.Text.Trim()+ "'  ";
                    sql += "  and c.receptid=receptid and c.productcode=productcode) and states='OK' group by receptid,productcode";
                    DataSet ds = DbAccess.SelectBySql(sql);
                    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        tranid = ds.Tables[0].Rows[0]["OKqty"].ToString();
                        item_id = ds.Tables[0].Rows[0]["ITEM_ID"].ToString();
                        org_id = ds.Tables[0].Rows[0]["org_id"].ToString();
                        srecept = ds.Tables[0].Rows[0]["receptid"].ToString();
                        productcode = ds.Tables[0].Rows[0]["productcode"].ToString();

                        txtenableqty.Text = ds.Tables[0].Rows[0]["OKqty"].ToString();
                        txtqty.Focus();
                        txtlotno.Enabled = false;
                    }
                    else
                    {
                        lblinfo.ForeColor = Color.Red;
                        lblinfo.Text = txtlotno.Text + "不存在或没有良品数";
                        txtlotno.Focus();
                        txtlotno.Text = "";
                        txtlotno.Enabled = true;
                    }
                }

            }
        }
        private void txtqty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
                //btnOK_Click(sender, null);
                txtreason.Focus();
        }

        private void txtreason_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
                btnOK_Click(sender, null);
        }


    }
}