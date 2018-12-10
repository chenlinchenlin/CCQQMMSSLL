using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;

namespace DX_QMS
{
    public partial class EMSOtherReceiveRpt : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public EMSOtherReceiveRpt()
        {
            InitializeComponent();
           
            this.txtstates.SelectedIndex = 0;
            //////// setRule();
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
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "客供报表");
            btndel.Enabled = bool.Parse(dic["hasDelete"].ToString());
            btnupdate.Enabled = bool.Parse(dic["hasUpdate"].ToString());
        }

        private void EMSOtherReceiveRpt_Load(object sender, EventArgs e)
        {
            txtarrivtedate1.EditValue = DateTime.Now.ToLocalTime().ToString();
            txtarrivatedate2.Text = "";
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }

        public static DataSet delivery_emsotherrecReport(string opertype, string receptid, string date1, string date2, string vendorcode, string vendorname, string materialcode, string materialname, string states, string remarks, string po)
        {
            SqlParameter[] para = new SqlParameter[11];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@receptid", receptid);
            para[2] = new SqlParameter("@date1", date1);
            para[3] = new SqlParameter("@date2", date2);
            para[4] = new SqlParameter("@vendorcode", vendorcode);
            para[5] = new SqlParameter("@vendorname", vendorname);
            para[6] = new SqlParameter("@materialcode", materialcode);
            para[7] = new SqlParameter("@materialname", materialname);
            para[8] = new SqlParameter("@states", states);
            para[9] = new SqlParameter("@remarks", remarks);
            para[10] = new SqlParameter("@po", po);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "Delivery_EMSOtherRecReport", para);
        }

        void selectdata(object obj)
        {

            databind.DataSource = null;
            string begindate = "", enddate = "";
            if (txtarrivtedate1.Text != "")
                begindate = txtarrivtedate1.EditValue.ToString();
            if (txtarrivatedate2.Text != "")
                enddate = txtarrivatedate2.EditValue.ToString();
            if (begindate == "" && this.txtcustomer.Text.Trim().Equals("") && txtproduct.Text.Equals("") && txtcusproduct.Text == "" && txtreceptid.Text == "")
            {
                MessageBox.Show("请至少输入两个条件！");
                return;
            }

            DataSet ds = delivery_emsotherrecReport("查询", txtreceptid.Text.Trim(), begindate, enddate, txtcusproduct.Text.Trim(), txtcustomer.Text.Trim(), txtproduct.Text.Trim(),
                                                                  txtprodes.Text.Trim(), txtstates.Text, txtremarks.Text.Trim(), txtpo.Text.Trim());

            string p = "", id = "";
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                p += "'" + ds.Tables[0].Rows[i]["HYT编码"].ToString() + "',";
                id += "'" + ds.Tables[0].Rows[i]["单号"].ToString() + "',";
            }

            string ERPInputStock = "select to_char(INVENTORY_ITEM_ID) INVENTORY_ITEM_ID,MAX(ITEM_NUM) ITEM_NUM,TRANSACTION_REFERENCE,NVL(sum(TRANSACTION_QUANTITY),0) TRANSACTION_QUANTITY,MAX(TRANSACTION_DATE) TRANSACTION_DATE from apps. CUX_MTL_TRANSACTIONS_V where TRANSACTION_DATE>=to_date('2014-01-01','YYYY-mm-DD')";
            ERPInputStock += " and length(TRANSACTION_REFERENCE)=10 and TRANSACTION_REFERENCE>='1601010001'  and TRANSACTION_REFERENCE not like '%-%' and  TRANSACTION_REFERENCE not like '%盘%' ";
            ERPInputStock += "  group by INVENTORY_ITEM_ID,TRANSACTION_REFERENCE ";
            string ERPArea = " select SUB_INVENTORY_POSTION,ITEM_CODE from CUX_INV0036B_TEMP where SUB_INVENTORY_CODE='1046669'";
            DataSet dsERP = DbAccess.SelectByOracle(ERPInputStock);
            DataSet dsERPArea = DbAccess.SelectByOracle(ERPArea);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                id = id.TrimEnd(',');
                p = p.TrimEnd(',');
                DataRow[] rwerp = dsERP.Tables[0].Select("ITEM_NUM in(" + p + ") and TRANSACTION_REFERENCE in(" + id + ")");
                DataTable dterp = null;
                dterp = dsERP.Tables[0].Clone();
                for (int m = 0; m < rwerp.Length; m++)
                {
                    dterp.Rows.Add(rwerp[m].ItemArray);
                }
                //查询ERP系统上的入库数量


                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    string ssql = "select 1 from delivery where deliveryid='" + ds.Tables[0].Rows[i]["单号"].ToString() + "' and materialcode='" + ds.Tables[0].Rows[i]["HYT编码"].ToString() + "'";
                    DataSet dssql = DbAccess.SelectBySql(ssql);
                    if (dssql == null || dssql.Tables.Count <= 0 || dssql.Tables[0].Rows.Count <= 0)
                        ds.Tables[0].Rows[i]["状态"] = "待生成标签";
                    else
                    {
                        for (int j = 0; j < dterp.Rows.Count; j++)
                        {
                            //if (ds.Tables[0].Rows[i]["id"].ToString() == dterp.Rows[j]["INVENTORY_ITEM_ID"].ToString() && ds.Tables[0].Rows[i]["单号"].ToString() == dterp.Rows[j]["TRANSACTION_REFERENCE"].ToString())
                            if (ds.Tables[0].Rows[i]["HYT编码"].ToString() == dterp.Rows[j]["ITEM_NUM"].ToString() && ds.Tables[0].Rows[i]["单号"].ToString() == dterp.Rows[j]["TRANSACTION_REFERENCE"].ToString())
                            {
                                int ERPQty = 0;
                                ERPQty = int.Parse(dterp.Rows[j]["TRANSACTION_QUANTITY"].ToString());
                                if ((int.Parse(ds.Tables[0].Rows[i]["OK数"].ToString()) + int.Parse(ds.Tables[0].Rows[i]["NG数"].ToString())) >= int.Parse(ds.Tables[0].Rows[i]["来料数"].ToString()) && ERPQty > 0 && ERPQty < int.Parse(ds.Tables[0].Rows[i]["OK数"].ToString()))
                                    ds.Tables[0].Rows[i]["状态"] = "正在入库";
                                else if (ERPQty > 0 && ERPQty >= int.Parse(ds.Tables[0].Rows[i]["OK数"].ToString()) && (int.Parse(ds.Tables[0].Rows[i]["OK数"].ToString()) + int.Parse(ds.Tables[0].Rows[i]["NG数"].ToString())) == int.Parse(ds.Tables[0].Rows[i]["来料数"].ToString()))
                                {
                                    ds.Tables[0].Rows[i]["状态"] = "入库完成";
                                    ds.Tables[0].Rows[i]["入库完成时间"] = dterp.Rows[j]["TRANSACTION_DATE"].ToString();
                                }
                                else if (ERPQty > 0 && (int.Parse(ds.Tables[0].Rows[i]["OK数"].ToString()) + int.Parse(ds.Tables[0].Rows[i]["NG数"].ToString())) < int.Parse(ds.Tables[0].Rows[i]["来料数"].ToString()))
                                {
                                    ds.Tables[0].Rows[i]["状态"] = "正在检验";
                                }
                                ds.Tables[0].Rows[i]["入库数"] = ERPQty.ToString();
                                break;
                            }
                        }
                        for (int j = 0; j < dsERPArea.Tables[0].Rows.Count; j++)
                        {
                            if (ds.Tables[0].Rows[i]["HYT编码"].ToString() == dsERPArea.Tables[0].Rows[j]["ITEM_CODE"].ToString())
                            {
                                ds.Tables[0].Rows[i]["库位"] = dsERPArea.Tables[0].Rows[j]["SUB_INVENTORY_POSTION"].ToString();
                                break;
                            }

                        }
                    }
                }

                if (!this.txtstates.Text.Equals("ALL"))
                {
                    DataRow[] arrDr = ds.Tables[0].Select("状态 <> '" + this.txtstates.Text + "'");
                    foreach (DataRow dr in arrDr)
                        ds.Tables[0].Rows.Remove(dr);
                }
                databind.DataSource = ds.Tables[0];



            }
            else
            {
                MessageBox.Show("没有符合条件的记录", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }      
        }
        private void btnsearch_Click(object sender, EventArgs e)
        {


            BackgroundTask.BackgroundWork(selectdata, null);

            // databind.DataSource = null;
            // string begindate = "", enddate = "";
            // if (txtarrivtedate1.Text !="")
            //     begindate = txtarrivtedate1.EditValue.ToString();
            // if (txtarrivatedate2.Text != "")
            //     enddate = txtarrivatedate2.EditValue.ToString();
            // if (begindate == "" && this.txtcustomer.Text.Trim().Equals("") && txtproduct.Text.Equals("") && txtcusproduct.Text == "" && txtreceptid.Text == "")
            // {
            //     MessageBox.Show("请至少输入两个条件！");
            //     return;
            // }

            // DataSet ds = delivery_emsotherrecReport("查询", txtreceptid.Text.Trim(), begindate, enddate, txtcusproduct.Text.Trim(), txtcustomer.Text.Trim(), txtproduct.Text.Trim(),
            //                                                       txtprodes.Text.Trim(), txtstates.Text, txtremarks.Text.Trim(), txtpo.Text.Trim());


            //string p = "", id = "";
            //for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            //{
            //    p += "'" + ds.Tables[0].Rows[i]["HYT编码"].ToString() + "',";
            //    id += "'" + ds.Tables[0].Rows[i]["单号"].ToString() + "',";
            //}

            //string ERPInputStock = "select to_char(INVENTORY_ITEM_ID) INVENTORY_ITEM_ID,MAX(ITEM_NUM) ITEM_NUM,TRANSACTION_REFERENCE,NVL(sum(TRANSACTION_QUANTITY),0) TRANSACTION_QUANTITY,MAX(TRANSACTION_DATE) TRANSACTION_DATE from apps. CUX_MTL_TRANSACTIONS_V where TRANSACTION_DATE>=to_date('2014-01-01','YYYY-mm-DD')";
            //ERPInputStock += " and length(TRANSACTION_REFERENCE)=10 and TRANSACTION_REFERENCE>='1601010001'  and TRANSACTION_REFERENCE not like '%-%' and  TRANSACTION_REFERENCE not like '%盘%' ";
            //ERPInputStock += "  group by INVENTORY_ITEM_ID,TRANSACTION_REFERENCE ";
            //string ERPArea = " select SUB_INVENTORY_POSTION,ITEM_CODE from CUX_INV0036B_TEMP where SUB_INVENTORY_CODE='1046669'";
            //DataSet dsERP = DbAccess.SelectByOracle(ERPInputStock);
            //DataSet dsERPArea = DbAccess.SelectByOracle(ERPArea);
            //if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            //{
            //    id = id.TrimEnd(',');
            //    p = p.TrimEnd(',');
            //    DataRow[] rwerp = dsERP.Tables[0].Select("ITEM_NUM in(" + p + ") and TRANSACTION_REFERENCE in(" + id + ")");
            //    DataTable dterp = null;
            //    dterp = dsERP.Tables[0].Clone();
            //    for (int m = 0; m < rwerp.Length; m++)
            //    {
            //        dterp.Rows.Add(rwerp[m].ItemArray);
            //    }
            //    //查询ERP系统上的入库数量


            //    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            //    {
            //        string ssql = "select 1 from delivery where deliveryid='" + ds.Tables[0].Rows[i]["单号"].ToString() + "' and materialcode='" + ds.Tables[0].Rows[i]["HYT编码"].ToString() + "'";
            //        DataSet dssql = DbAccess.SelectBySql(ssql);
            //        if (dssql == null || dssql.Tables.Count <= 0 || dssql.Tables[0].Rows.Count <= 0)
            //            ds.Tables[0].Rows[i]["状态"] = "待生成标签";
            //        else
            //        {
            //            for (int j = 0; j < dterp.Rows.Count; j++)
            //            {
            //                //if (ds.Tables[0].Rows[i]["id"].ToString() == dterp.Rows[j]["INVENTORY_ITEM_ID"].ToString() && ds.Tables[0].Rows[i]["单号"].ToString() == dterp.Rows[j]["TRANSACTION_REFERENCE"].ToString())
            //                if (ds.Tables[0].Rows[i]["HYT编码"].ToString() == dterp.Rows[j]["ITEM_NUM"].ToString() && ds.Tables[0].Rows[i]["单号"].ToString() == dterp.Rows[j]["TRANSACTION_REFERENCE"].ToString())
            //                {
            //                    int ERPQty = 0;
            //                    ERPQty = int.Parse(dterp.Rows[j]["TRANSACTION_QUANTITY"].ToString());
            //                    if ((int.Parse(ds.Tables[0].Rows[i]["OK数"].ToString()) + int.Parse(ds.Tables[0].Rows[i]["NG数"].ToString())) >= int.Parse(ds.Tables[0].Rows[i]["来料数"].ToString()) && ERPQty > 0 && ERPQty < int.Parse(ds.Tables[0].Rows[i]["OK数"].ToString()))
            //                        ds.Tables[0].Rows[i]["状态"] = "正在入库";
            //                    else if (ERPQty > 0 && ERPQty >= int.Parse(ds.Tables[0].Rows[i]["OK数"].ToString()) && (int.Parse(ds.Tables[0].Rows[i]["OK数"].ToString()) + int.Parse(ds.Tables[0].Rows[i]["NG数"].ToString())) == int.Parse(ds.Tables[0].Rows[i]["来料数"].ToString()))
            //                    {
            //                        ds.Tables[0].Rows[i]["状态"] = "入库完成";
            //                        ds.Tables[0].Rows[i]["入库完成时间"] = dterp.Rows[j]["TRANSACTION_DATE"].ToString();
            //                    }
            //                    else if (ERPQty > 0 && (int.Parse(ds.Tables[0].Rows[i]["OK数"].ToString()) + int.Parse(ds.Tables[0].Rows[i]["NG数"].ToString())) < int.Parse(ds.Tables[0].Rows[i]["来料数"].ToString()))
            //                    {
            //                        ds.Tables[0].Rows[i]["状态"] = "正在检验";
            //                    }
            //                    ds.Tables[0].Rows[i]["入库数"] = ERPQty.ToString();
            //                    break;
            //                }
            //            }
            //            for (int j = 0; j < dsERPArea.Tables[0].Rows.Count; j++)
            //            {
            //                if (ds.Tables[0].Rows[i]["HYT编码"].ToString() == dsERPArea.Tables[0].Rows[j]["ITEM_CODE"].ToString())
            //                {
            //                    ds.Tables[0].Rows[i]["库位"] = dsERPArea.Tables[0].Rows[j]["SUB_INVENTORY_POSTION"].ToString();
            //                    break;
            //                }

            //            }
            //        }
            //    }

            //    if (!this.txtstates.Text.Equals("ALL"))
            //    {
            //        DataRow[] arrDr = ds.Tables[0].Select("状态 <> '" + this.txtstates.Text + "'");
            //        foreach (DataRow dr in arrDr)
            //            ds.Tables[0].Rows.Remove(dr);
            //    }
            //    databind.DataSource = ds.Tables[0];



            //}
            //else
            //{
            //    MessageBox.Show("没有符合条件的记录", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}




            // /*
            // foreach (DataGridViewColumn dgvc in databind.Columns)
            // {
            //     if (dgvc.Name != "备注")
            //     {
            //         dgvc.ReadOnly = true;
            //     }
            //     else
            //     {
            //         dgvc.ReadOnly = false;
            //     }
            // }
            // */


        }

        public static string delivery_emsotherrec(string opertype, string receptid, string arrivatedate, string vendorcode, string vendorname, string materialcode, string materialname, int qty, int reelidqty,
                                                string states, string remarks, string po, string userid, string customer)
        {
            SqlParameter[] para = new SqlParameter[15];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@receptid", receptid);
            para[2] = new SqlParameter("@arrivatedate", arrivatedate);
            para[3] = new SqlParameter("@vendorcode", vendorcode);
            para[4] = new SqlParameter("@vendorname", vendorname);
            para[5] = new SqlParameter("@materialcode", materialcode);
            para[6] = new SqlParameter("@materialname", materialname);
            para[7] = new SqlParameter("@qty", qty);
            para[8] = new SqlParameter("@reelidqty", reelidqty);
            para[9] = new SqlParameter("@states", states);
            para[10] = new SqlParameter("@remarks", remarks);
            para[11] = new SqlParameter("@po", po);
            para[12] = new SqlParameter("@userid", userid);
            para[13] = new SqlParameter("@customer", customer);
            para[14] = new SqlParameter("@msg", SqlDbType.VarChar, 200);
            para[14].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "Delivery_EMSotherImport", para);
            return para[14].Value.ToString();
        }

        private void btndel_Click(object sender, EventArgs e)
        {
            int  n= 0 ;
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length < 1)
                return;

            for (int k = gridView.GetSelectedRows().Length; k > 0; k--)
            {

                DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[k - 1]);
                string msgdel = delivery_emsotherrec("删除", dr["单号"].ToString(), "", dr["客户编码"].ToString(), "", dr["HYT编码"].ToString(), "", 0,
                     0, "", "", "", Login.username,"");

                if (msgdel.IndexOf("成功") >= 0)
                {
                    n = n + 1;
                    gridView.DeleteRow(gridView.GetSelectedRows()[k - 1]);

                }
            }
        }

        private void btnupdate_Click(object sender, EventArgs e)
        {
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length < 1)
                return;
            int  n=0;

            for (int k = gridView.GetSelectedRows().Length; k > 0; k--)
            {

                DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[k - 1]);
                string msgdel = delivery_emsotherrec("更新", dr["单号"].ToString(), "", dr["客户编码"].ToString(), "", dr["HYT编码"].ToString(), "", 0,
                     0, "", dr["备注"].ToString(), "", Login.username, "");
                if (msgdel.IndexOf("成功") >= 0)
                {
                    n = n + 1;
                }

            }
            MessageBox.Show("总共有:" + n.ToString() + "项更新成功", "提示");
        }

        private void gridView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {

            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            if (gridView.FocusedRowHandle < 0)
                return;

            try
            {
                if (e.Column.FieldName == "OK数" && int.Parse(gridView.GetFocusedRowCellValue("OK数").ToString()) > 0)
                {
                    EMSOtherRecList EMSL = new EMSOtherRecList(gridView.GetFocusedRowCellValue("单号").ToString(), gridView.GetFocusedRowCellValue("HYT编码").ToString(), "OK明细", "0");
                    EMSL.Show();
                }
                else if (e.Column.FieldName == "NG数" && int.Parse(gridView.GetFocusedRowCellValue("NG数").ToString()) > 0)
                {
                    EMSOtherRecList EMSL = new EMSOtherRecList(gridView.GetFocusedRowCellValue("单号").ToString(), gridView.GetFocusedRowCellValue("HYT编码").ToString(), "NG明细", "0");
                    EMSL.Show();
                }
                else if (e.Column.FieldName == "入库数" && int.Parse(gridView.GetFocusedRowCellValue("入库数").ToString()) > 0)
                {
                    EMSOtherRecList EMSL = new EMSOtherRecList(gridView.GetFocusedRowCellValue("单号").ToString(), gridView.GetFocusedRowCellValue("HYT编码").ToString(), "入库明细", gridView.GetFocusedRowCellValue("id").ToString());
                    EMSL.Show();
                }

            }
            catch
            {

            }

            }

        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "客供报表信息";
            int n = name.LastIndexOf(".") + 1;
            if (n > 0) name = name.Substring(n, name.Length - n);
            dlg.Title = "Export To " + title;
            dlg.FileName = name;
            dlg.Filter = filter;
            if (dlg.ShowDialog() == DialogResult.OK) return dlg.FileName;
            return "";
        }
        private void ExportToEx(String filename, string ext, BaseView exportView)
        {
            Cursor currentCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            if (ext == "rtf") exportView.ExportToRtf(filename);
            if (ext == "pdf") exportView.ExportToPdf(filename);
            if (ext == "mht") exportView.ExportToMht(filename);
            if (ext == "htm") exportView.ExportToHtml(filename);
            if (ext == "txt") exportView.ExportToText(filename);
            if (ext == "xls") exportView.ExportToXls(filename);
            if (ext == "xlsx") exportView.ExportToXlsx(filename);
            Cursor.Current = currentCursor;
        }
        private void OpenFile(string fileName)
        {
            if (XtraMessageBox.Show("Do you want to open this file?", "Export To...", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    System.Diagnostics.Process process = new System.Diagnostics.Process();
                    process.StartInfo.FileName = fileName;
                    process.StartInfo.Verb = "Open";
                    process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
                    process.Start();
                }
                catch
                {
                    DevExpress.XtraEditors.XtraMessageBox.Show(this, "Cannot find an application on your system suitable for openning the file with exported data.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            progressBarControl1.Position = 0;
        }
        void gridView_PrintExportProgress(object sender, ProgressChangedEventArgs e)
        {
            SetPosition(e.ProgressPercentage);
        }
        void SetPosition(int pos)
        {
            progressBarControl1.Position = pos;
            this.Update();
        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            DataTable dt = databind.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0) return;

            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xls");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xls", gridView);
            OpenFile(fileName);
        }

        private void sBtnreset_Click(object sender, EventArgs e)
        {
            txtarrivtedate1.Text = "";
            txtarrivatedate2.Text = "";
            txtcusproduct.Text = "";
            txtcustomer.Text = "";
            txtproduct.Text = "";
            txtpo.Text = "";
            txtstates.Text = "";
            txtreceptid.Text = "";
            txtprodes.Text = "";
            txtremarks.Text = "";
            databind.DataSource = null;
        }

    }
}