using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Windows.Forms;
using DevExpress.XtraBars;
using System.Data;
using System.Data.SqlClient;
using DX_QMS.Common;

namespace DX_QMS
{
    public partial class MaterialLot : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public MaterialLot()
        {
            InitializeComponent();         
        }

        private void bindorg_id()
        {
            string sql = "select ORG_ID,ORG_NAME from ORG_INFO where ORG_TYPE='ORG' order by SORT ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtorg_id.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txtorg_id.Properties.Items.Add(row["ORG_NAME"]);
            }
            txtorg_id.SelectedIndex = 0;
        }

        private void ChangeState(string sType)
        {
            if (sType.Equals("ByDate"))
            {
                this.tbDeliveryID.Text = "";
                this.tbEndDeliveryID.Text = "";
                this.tbDeliveryID.Enabled = false;
                this.tbEndDeliveryID.Enabled = false;

                this.dtpEndDate.Enabled = true;
                this.dtpStartDate.Enabled = true;
                this.numUDEndHour.Enabled = true;
                this.numUDStartHour.Enabled = true;
                this.numUDEndHour.Value = DateTime.Now.Hour;
                this.tbERPUserID.Focus();
            }
            if (sType.Equals("ByDelivery"))
            {
                this.tbERPUserID.Text = "";
                this.dtpEndDate.Enabled = false;
                this.dtpStartDate.Enabled = false;
                this.numUDEndHour.Enabled = false;
                this.numUDStartHour.Enabled = false;

                this.tbDeliveryID.Enabled = true;
                this.tbEndDeliveryID.Enabled = true;
                this.tbDeliveryID.Focus();
                this.tbDeliveryID.SelectAll();
            }
        }

        private void selecttype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (selecttype.SelectedIndex == 0)
            {
                ChangeState("ByDelivery");
            }
            else
            {
                ChangeState("ByDate");
            }
        }

        private void MaterialLot_Load(object sender, EventArgs e)
        {
            bindorg_id();
            ChangeState("ByDelivery");
            selecttype.SelectedIndex = 0;
        }

        private void tbDeliveryID_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && tbDeliveryID.Text !="")
            {
                this.btnAccept.Focus();
                btnAccept_Click(sender, e);
            }
        }

        private void tbEndDeliveryID_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && tbEndDeliveryID.Text != "")
            {
                this.btnAccept.Focus();
                btnAccept_Click(sender, e);
            }
        }

        private void btnAccept_Click(object sender, EventArgs e)
        {
            if (this.btnAccept.Enabled == false)
                return;
            this.dGVLots.DataSource = null;
            this.dGVDelivery.DataSource = null;
         
            string sOraSql = " select receipt_num, max(transaction_id) transactionid, round(sum(quantity),0) qty, item_number materialcode, max(item_desc) materialname, max(user_name) user_name,";
            sOraSql += " to_char(max(transaction_date),'yyyy-mm-dd hh24:mi:ss') transaction_date, max(unit_of_measure) unit, max(vendor_name) vendorname, max(vendor_number) vendorcode,max(ORGANIZATION_ID) org_id,max(lot_number) lot_number ";
            sOraSql += " from cux_inv_rec3_v ";

            if (this.selecttype.SelectedIndex == 0)
            {
                string sSDeliveryID = this.tbDeliveryID.Text.Trim();
                string sEDeliveryID = this.tbEndDeliveryID.Text.Trim();
                sOraSql += " where transaction_type in ('TRANSFER','RECEIVE','ACCEPT') and ORGANIZATION_CODE='" + txtorg_id.Text.Trim() + "'";
                if (sSDeliveryID.Equals("") && sEDeliveryID.Equals(""))
                {
                    MessageBox.Show("尚未输入接收单号", "提示");
                    return;
                }
                else if (!sSDeliveryID.Equals("") && !sEDeliveryID.Equals(""))
                {
                    sOraSql += " and receipt_num >=" + sSDeliveryID + " and receipt_num <=" + sEDeliveryID + "";
                }
                else
                {
                    if (!sSDeliveryID.Equals(""))
                        sOraSql += " and receipt_num =" + sSDeliveryID + "";
                    if (!sEDeliveryID.Equals(""))
                        sOraSql += " and receipt_num =" + sEDeliveryID + "";
                }
            }
            else
            {
                string sSDate = dtpStartDate.Text + " " + this.numUDStartHour.Value.ToString() + ":00:00";
                string sEDate = dtpEndDate.Text + " " + this.numUDEndHour.Value.ToString() + ":59:59";
                sOraSql += " where transaction_type in ('TRANSFER','RECEIVE','ACCEPT') ";
                sOraSql += " and transaction_date>=to_date('" + sSDate + "','yyyy/mm/dd hh24:mi:ss')";
                sOraSql += " and transaction_date<=to_date('" + sEDate + "','yyyy/mm/dd hh24:mi:ss')";

            }
            string sERPUserID = this.tbERPUserID.Text.Trim();
            if (!sERPUserID.Equals(""))
                sOraSql += " and vendor_number='" + sERPUserID + "'";

            sOraSql += " group by receipt_num, item_number";
            DataSet dsAgain = DbAccess.SelectByOracle(sOraSql);

            if (dsAgain.Tables.Count > 0 && dsAgain.Tables[0].Rows.Count > 0)
            {
                DataTable dt = dsAgain.Tables[0];
                for (int i = dt.Rows.Count - 1; i > 0; i--)
                {
                    string sql = "select 1 from delivery where deliveryid='" + dt.Rows[i]["receipt_num"].ToString() + "' and materialcode='" + dt.Rows[i]["materialcode"].ToString() + "'";
                    DataTable dtdelivery = DbAccess.SelectBySql(sql).Tables[0];
                    if (dtdelivery.Rows.Count > 0)
                    {
                        dt.Rows.RemoveAt(i);
                    }
                }
                if (dt.Rows.Count > 0)
                {
                    this.dGVDelivery.DataSource = dt;
                    //this.dGVDelivery.Focus();
                }
            }
            else
            {
                MessageBox.Show("未找到符合条件的接收单号", "提示");
                return;
            }
        }

        private void btnclear_Click(object sender, EventArgs e)
        {
            txtorg_id.Text = "";
            tbDeliveryID.Text = "";
            tbEndDeliveryID.Text = "";
            dtpStartDate.Text = "";
            numUDStartHour.Text = "0";
            dtpEndDate.Text = "";
            numUDEndHour.Text = "0";
            tbERPUserID.Text = "";
            selecttype.SelectedIndex = 0;
            dGVDelivery.DataSource = null;
            dGVLots.DataSource = null;

        }

        private void btnsuppliersearch_Click(object sender, EventArgs e)
        {

        }

        private void dGVDeliverygridview_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            try
            {
                DataTable dt = dGVDelivery.DataSource as DataTable;
                if (dt == null || dt.Rows.Count < 1)
                {
                    return;
                }

                string sql = "select 1 from UserMaterial where MaterialId='" + dGVDeliverygridview.GetDataRow(e.RowHandle)["materialcode"].ToString() + "'";
                DataTable ds = DbAccess.SelectBySql(sql).Tables[0];
                if (ds != null && ds.Rows.Count > 0)
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
                if (dGVDeliverygridview.GetDataRow(e.RowHandle)["materialname"].ToString().Contains("专用】"))
                {
                    e.Appearance.BackColor = Color.LightSalmon;
                }
                else
                {
                    e.Appearance.BackColor = Color.White;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dGVLotsgridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            DataTable dt = dGVLots.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            try
            {
                if (dGVLotsgridView.GetDataRow(e.RowHandle)["printflag"].ToString() == "1")
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
                else
                {
                    e.Appearance.BackColor = Color.White;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            dGVLots.DataSource = null;
        }





        public static string Delivery_Add(string sDeliveryID, string sItem, string sMaterialCode, string sMaterialName, string sLotNo, int iQty,
                    string sInDate, string sVendorCode, string sVendorName, string sUserID, string cartron, string orgid, string pindian)
        {
            SqlParameter[] para = new SqlParameter[14];
            para[0] = new SqlParameter("@deliveryid", sDeliveryID);
            para[1] = new SqlParameter("@item", sItem);
            para[2] = new SqlParameter("@materialcode", sMaterialCode);
            para[3] = new SqlParameter("@materialname", sMaterialName);
            para[4] = new SqlParameter("@lotno", sLotNo);
            para[5] = new SqlParameter("@qty", iQty);
            para[6] = new SqlParameter("@indate", sInDate);
            para[7] = new SqlParameter("@vendorcode", sVendorCode);
            para[8] = new SqlParameter("@vendorname", sVendorName);
            para[9] = new SqlParameter("@userid", sUserID);
            para[10] = new SqlParameter("@carton", cartron);
            para[11] = new SqlParameter("@orgid", orgid);
            para[12] = new SqlParameter("@lot_number", pindian);
            para[13] = new SqlParameter("@errormsg", SqlDbType.VarChar, 200);
            para[13].Direction = ParameterDirection.Output;

            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "delivery_add_hk", para);
            return para[13].Value.ToString();
        }
        private DataTable dataHK;

        private void btnGen_Click(object sender, EventArgs e)
        {

            try
            {
                if (this.btnGen.Enabled == false) return;
                this.dGVLots.DataSource = null;
                if (dGVDeliverygridview.GetSelectedRows().Length > 0)
                {
                    //修改为不同的接收单同时生成
                    string sDeliveryIDs = "";
                    for (int i = 0; i < dGVDeliverygridview.GetSelectedRows().Length; i++)
                    {

                        DataRow dr = dGVDeliverygridview.GetDataRow(dGVDeliverygridview.GetSelectedRows()[i]);

                        if (dr["materialname"].ToString().Contains("专用】"))
                        {
                            if (MessageBox.Show("系统判定该物料应该做Reel ID,请确认是否要继续?", "Reel ID提示", MessageBoxButtons.YesNo) == DialogResult.No)
                                continue;
                        }
                        if (dr["org_id"].ToString() == "HCL")
                        {
                            if (MessageBox.Show("HCL组织的接收单号不能在这里生成批次号！", "Reel ID提示", MessageBoxButtons.YesNo) == DialogResult.No)
                                return;
                        }
                        //2013年6月17号新增防爆物料生成ReelID,是防爆料不生成批次(要到电子料界面去生成ReelID)

                  
                        string sql = "select 1 from UserMaterial where MaterialId='" + dr["materialcode"].ToString() + "'";
                        DataTable dst = DbAccess.SelectBySql(sql).Tables[0];
                        if (dst != null && dst.Rows.Count > 0)
                        {
                            continue;
                        }

                        string spindian = dr["lot_number"].ToString();
                        string sorgid = dr["org_id"].ToString();
                        string sDeliveryID = dr["receipt_num"].ToString();
                        string sItem = dr["transactionid"].ToString();
                        string sMaterialCode = dr["materialcode"].ToString();
                        string sMaterialName = dr["materialname"].ToString();
                        int iQty = Convert.ToInt32(dr["qty"].ToString());
                        string sInDate = (dr["transaction_date"].ToString() == "") ? "" : dr["transaction_date"].ToString();
                        string sVendorCode = (dr["vendorcode"].ToString() == "") ? "" : dr["vendorcode"].ToString();
                        string sVendorName = (dr["vendorname"].ToString() == "") ? "" : dr["vendorname"].ToString();
                        sDeliveryIDs += "'" + sDeliveryID + "',";

                        string carton = "";
                        if (this.dataHK != null)
                        {
                            DataRow[] rows = this.dataHK.Select("接收号 ='" + sDeliveryID + "' and 物料编码='" + sMaterialCode + "'");
                            if (rows.Length > 0)
                                carton = rows[0][0].ToString().Substring(0, rows[0][0].ToString().IndexOf('-') + 1);
                            foreach (DataRow row in rows)
                            {
                                carton = carton + row[0].ToString().Substring(row[0].ToString().IndexOf('-') + 1) + "/";
                            }
                            carton = carton.TrimEnd('/');
                        }
                        //sDeliveryIDs = sDeliveryID;
                       Delivery_Add(sDeliveryID, sItem, sMaterialCode, sMaterialName, "", iQty, sInDate, sVendorCode, sVendorName, Login.userId, carton, sorgid, spindian);
                    }
                    //if (this.dataHK != null)
                    //    this.dataHK = null;
                    sDeliveryIDs = sDeliveryIDs.TrimEnd(',');
                    string sSql = " select d.lotno,d.deliveryid,d.item,d.materialcode,s.materialname,d.qty,case d.state when 'Used' then '已入库' else '未入库' end as state,d.transactiondate,d.eventuser,d.eventtime,";
                    sSql += " 'printqty'=case when isnull(s.inminpack,1)=1 then 10 when d.qty%s.inminpack=0 then d.qty/s.inminpack else d.qty/s.inminpack+1 end, isnull(s.inminpack,1) minpack, s.materialenname,d.printflag,d.carton ";
                    sSql += " from delivery_v d left join materialspec s on s.materialcode=d.materialcode and s.organization='84'";
                    sSql += " where d.deliveryid in (" + sDeliveryIDs + ") ";
                    //sSql += " where d.funidx like ('" + sDeliveryIDs + "%') ";
                    DataSet ds = DbAccess.SelectBySql(sSql);
                    if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        this.dGVLots.DataSource = ds.Tables[0];
                        this.dGVLots.Focus();
                    }
                    else
                    {
                        MessageBox.Show("输入的接收单号不存在，或者已入库");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            if (this.btnSelect.Enabled == false)
                return;
           
            this.dGVLots.DataSource = null;
            string sSql = " select d.lotno,d.deliveryid,d.item,d.materialcode,s.materialname,d.qty, case d.state when 'Used' then '已入库' else '未入库' end as state,d.transactiondate,d.eventuser,case when d.lotno like 'Z%' then pushtimestamp else d.eventtime end eventtime,";
            sSql += " 'printqty'=case when isnull(s.inminpack,1)=1 then 10 when d.qty%s.inminpack=0 then d.qty/s.inminpack else d.qty/s.inminpack+1 end, isnull(s.inminpack,1) minpack,s.materialenname,d.printflag,d.carton ";
            sSql += " from delivery d left join materialspec s on s.materialcode=d.materialcode and s.organization='84'";
            sSql += " where d.state='NotUse'";
            if (selecttype.SelectedIndex == 0)
            {
                string sSDeliveryID = this.tbDeliveryID.Text.Trim();
                string sEDeliveryID = this.tbEndDeliveryID.Text.Trim();
                if (!sSDeliveryID.Equals("") && !sEDeliveryID.Equals(""))
                {
                    sSql += " and d.deliveryid >='" + sSDeliveryID + "' and d.deliveryid <='" + sEDeliveryID + "'";
                }
                else
                {
                    if (!sSDeliveryID.Equals(""))
                        sSql += " and d.deliveryid ='" + sSDeliveryID + "'";
                    if (!sEDeliveryID.Equals(""))
                        sSql += " and d.deliveryid ='" + sEDeliveryID + "'";
                }
                DataSet ds = DbAccess.SelectBySql(sSql);
                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    this.dGVLots.DataSource = ds.Tables[0];
            }
            else
            {
                string sSDate = dtpStartDate.Text + " " + this.numUDStartHour.Value.ToString() + ":00:00";
                string sEDate = dtpEndDate.Text + " " + this.numUDEndHour.Value.ToString() + ":59:59";
                sSql += " and transactiondate>='" + sSDate + "'";
                sSql += " and transactiondate<='" + sEDate + "'";
                string sERPUserID = this.tbERPUserID.Text.Trim();
                if (!sERPUserID.Equals(""))
                    sSql += " and vendorcode='" + sERPUserID + "'";

                DataSet ds = DbAccess.SelectBySql(sSql);
                if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    this.dGVLots.DataSource = ds.Tables[0];
            }
        }

 
    }
}