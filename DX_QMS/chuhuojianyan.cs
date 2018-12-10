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
using System.Data;
using System.Data.SqlClient;
using System.IO;
using DX_QMS.Common;

namespace DX_QMS
{
    public partial class chuhuojianyan : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Microsoft.Office.Interop.Excel.ApplicationClass appClsExcel = null;
        IQC ic = new IQC();
        private string ifscan = "否";
        private string lotno = "";
        public chuhuojianyan()
        {
            InitializeComponent();
            bindTestAQL();
            bindOrg();
            DataRow[] rwc;
            DataTable dtcus = bindtypes();
            DataTable dtcusnew = dtcus.Clone();
            rwc = dtcus.Select("Definetype='客户'");
            for (int i = 0; i < rwc.Length; i++)
            {
                dtcusnew.Rows.Add(rwc[i].ItemArray);
            }
            txtcustomer.DataSource = dtcusnew;
            txtcustomer.DisplayMember = dtcusnew.Columns[1].ToString();
            txtcustomer.ValueMember = dtcusnew.Columns[1].ToString();
            DataTable dtitem = dtcus.Clone();
            rwc = dtcus.Select("Definetype='测试项目'");
            for (int i = 0; i < rwc.Length; i++)
            {
                dtitem.Rows.Add(rwc[i].ItemArray);
            }
            txttestitem.DataSource = dtitem;
            txttestitem.DisplayMember = dtitem.Columns[1].ToString();
            txttestitem.ValueMember = dtitem.Columns[1].ToString();

            DataTable dttools = dtcus.Clone();
            rwc = dtcus.Select("Definetype='检验工具'");
            for (int i = 0; i < rwc.Length; i++)
            {
                dttools.Rows.Add(rwc[i].ItemArray);
            }
            txttesttools.DataSource = dttools;
            txttesttools.DisplayMember = dttools.Columns[1].ToString();
            txttesttools.ValueMember = dttools.Columns[1].ToString();
            DataTable dtline = dtcus.Clone();
            rwc = dtcus.Select("Definetype='线体'");
            for (int i = 0; i < rwc.Length; i++)
            {
                dtline.Rows.Add(rwc[i].ItemArray);
            }
            txtline.DataSource = dtline;
            txtline.DisplayMember = dtline.Columns[1].ToString();
            txtline.ValueMember = dtline.Columns[1].ToString();
            DataTable dtQE = dtcus.Clone();
            rwc = dtcus.Select("Definetype='责任QE'");
            for (int i = 0; i < rwc.Length; i++)
            {
                dtQE.Rows.Add(rwc[i].ItemArray);
            }
            txtQE.DataSource = dtQE;
            txtQE.DisplayMember = dtQE.Columns[1].ToString();
            txtQE.ValueMember = dtQE.Columns[1].ToString();
            DataTable dtmaster = dtcus.Clone();
            rwc = dtcus.Select("Definetype='责任主管'");
            for (int i = 0; i < rwc.Length; i++)
            {
                dtmaster.Rows.Add(rwc[i].ItemArray);
            }
            txtmaster.DataSource = dtmaster;
            txtmaster.DisplayMember = dtmaster.Columns[1].ToString();
            txtmaster.ValueMember = dtmaster.Columns[1].ToString();

            DataTable dtstate = dtcus.Clone();
            rwc = dtcus.Select("Definetype='产品状态'");
            for (int i = 0; i < rwc.Length; i++)
            {
                dtstate.Rows.Add(rwc[i].ItemArray);
            }
            txtstate.DataSource = dtstate;
            txtstate.DisplayMember = dtstate.Columns[1].ToString();
            txtstate.ValueMember = dtstate.Columns[1].ToString();
            txtIQCChecktype.SelectedIndex = 1;
            txtAQL.SelectedIndex = 1;
            setRule();
        }
        string name = "OQCTestList";
        private void setRule()
        {
            Dictionary<string, bool> dic = GroupPermission.SelectRulesForForm(name, Login.groupId);
            this.btnsave.Enabled = dic["hasInsert"]; 
            this.btndel.Enabled = dic["hasDelete"];

        }
        private void bindTestAQL()
        {
            DataTable dt = Common.DbAccess.SelectBySql("select AQL,AQLValue from IQC_TestAQL").Tables[0];
            txtAQL.DataSource = dt;
            txtAQL.DisplayMember = dt.Columns["AQL"].ToString();
            txtAQL.ValueMember = dt.Columns["AQLValue"].ToString();
        }
        public static DataSet SAP_GetORG(string type)
        {
            SqlParameter[] para = new SqlParameter[1];
            para[0] = new SqlParameter("@type", type);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "SAP_GetOrg", para);
        }
        private void bindOrg()
        {
            this.txtorg_id.DataSource = SAP_GetORG("ORG").Tables[0];
            this.txtorg_id.ValueMember = "ORG_ID";
            this.txtorg_id.DisplayMember = "ORG_NAME";
        }
        private DataTable bindtypes()
        {
            string ssql = "select Definetype,Definevalue from OQC_TypeDefine order by sort ";
            DataTable dt = Common.DbAccess.SelectBySql(ssql).Tables[0];
            return dt;
        }
        private string getconment(string testitem)
        {
            string s = "";
            string sql = "select code from OQC_TypeDefine where Definevalue='" + testitem + "'";
            DataTable dt = Common.DbAccess.SelectBySql(sql).Tables[0];
            if (dt.Rows.Count > 0)
                s = dt.Rows[0]["code"].ToString();
            return s;
        }
        private void chuhuojianyan_Load(object sender, EventArgs e)
        {
            txtworkno.Focus();
            txtconment.Text = getconment(txttestitem.SelectedValue.ToString());
        }
        private void txtworkno_Leave(object sender, EventArgs e)
        {
            databind.DataSource = null;
            if (txtworkno.Text == "") return;
            txtworkno.Leave -= txtworkno_Leave;
            //string sql = "select item_num,item_desc,start_quantity,job_desc,WO.DEPARTMENT_ID,completion_subinventory,order_number,gn_mo,V.ORGANIZATION_ID from cux_wip_header_v v,CUX_WIP_OPERATIONS_V WO  where " +
            //                      " WO.WIP_ENTITY_ID =v.WIP_ENTITY_ID and wip_entity_name =" + "'" + txtworkno.Text.ToUpper() + "' and WO.ORGANIZATION_ID = '" + txtorg_id.SelectedValue.ToString() + "'";
            string sql = "select item_num,item_desc,start_quantity,job_desc,WO.DEPARTMENT_ID,completion_subinventory,order_number,gn_mo,V.ORGANIZATION_ID from cux_wip_header_v v,CUX_WIP_OPERATIONS_V WO  where " +
                                  " WO.WIP_ENTITY_ID =v.WIP_ENTITY_ID and wip_entity_name =" + "'" + txtworkno.Text.ToUpper() + "' and v.ORGANIZATION_CODE = '" + txtorg_id.SelectedValue.ToString() + "'";
            DataSet ds = Common.DbAccess.SelectByOracle(sql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                lblinfo.Text = ds.Tables[0].Rows[0]["item_desc"].ToString();
                lblinfo.ForeColor = Color.Blue;
                txtproductcode.Text = ds.Tables[0].Rows[0]["item_num"].ToString();
                txtworkqty.Text = ds.Tables[0].Rows[0]["start_quantity"].ToString();

                if (txtworkno.Text.ToUpper().EndsWith("WS") || txtworkno.Text.ToUpper().EndsWith("AS") || txtworkno.Text.ToUpper().EndsWith("WN") || txtworkno.Text.ToUpper().EndsWith("AN"))
                {
                    string s = "";
                    string ssql = "select Definevalue,sort from OQC_TypeDefine where Definetype='检验项目' order by sort ";
                    DataTable dt = Common.DbAccess.SelectBySql(ssql).Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            s += dt.Rows[i]["sort"].ToString() + "." + dt.Rows[i]["Definevalue"].ToString() + " ";
                        }
                    }
                    DialogResult result = MessageBox.Show(s, "请确认是否检验完下面的项目!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (DialogResult.No == result)
                    {
                        txtworkno.Text = "";
                        txtworkno.Focus();
                        txtworkno.Leave += txtworkno_Leave;
                        return;
                    }
                }

                System.Collections.ArrayList ltL = new System.Collections.ArrayList();
                int index = 0;
                string str = ds.Tables[0].Rows[0]["item_desc"].ToString();
                foreach (Char ch in str)
                {
                    if (ch == '(')
                    {
                        ltL.Add(index);
                    }
                    index++;
                }

                System.Collections.ArrayList ltR = new System.Collections.ArrayList();
                int indexR = 0;
                foreach (Char ch in str)
                {
                    if (ch == ')')
                    {
                        ltR.Add(indexR);
                    }
                    indexR++;
                }
                try
                {
                    txtcuscode.Text = str.Substring(int.Parse(ltL[0].ToString()) + 1, int.Parse(ltR[0].ToString()) - int.Parse(ltL[0].ToString()) - 1);
                }
                catch
                {
                }
            }
            else
            {
                lblinfo.Text = txtworkno.Text + "该工单号不存在";
                lblinfo.ForeColor = Color.Red;
                txtworkno.Text = "";
                txtworkno.Focus();
            }
            txtworkno.Leave += txtworkno_Leave;
        }
        private void txttestitem_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txttestitem.SelectedValue == null) return;
            txtconment.Text = getconment(txttestitem.SelectedValue.ToString());
        }
        private void txtsamplebar_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && this.txtsamplebar.Text != "")
            {
                DataTable IfRelation = Common.DbAccess.SelectBySql("select checkdate,org_id, workno,serialno from OQC_SampleFactList where serialno='"
                                       + this.txtsamplebar.Text.Trim() + "'").Tables[0];
                if (IfRelation.Rows.Count > 0)
                {
                    if (IfRelation.Rows[0]["checkdate"].ToString() != "")
                    {
                        this.lblinfo.Text = "你输入的条码号:" + txtsamplebar.Text.ToUpper().Trim() + ",已装入工单:" + IfRelation.Rows[0]["workno"].ToString();
                        lblinfo.ForeColor = Color.Red;
                        txtsamplebar.Text = "";
                        txtsamplebar.Focus();
                        return;
                    }
                }

                for (int i = 0; i < datasamplelist.Rows.Count; i++)
                {
                    if (datasamplelist.Rows[i].Cells["SN"].Value.ToString() == txtsamplebar.Text.ToUpper())
                    {

                        this.lblinfo.Text = "你输入的条码号：" + txtsamplebar.Text + "已经输入";
                        lblinfo.ForeColor = Color.Red;
                        txtsamplebar.Text = "";
                        txtsamplebar.Focus();
                        return;
                    }
                }
                if (int.Parse(this.txtfactsampleqty.Text.Trim()) >= int.Parse(txtsampleqty.Text))
                {
                    this.lblinfo.Text = "已经输入完成,不能超过抽样数量:" + txtsampleqty.Text;
                    lblinfo.ForeColor = Color.Red;
                    this.txtsamplebar.Text = "";
                    return;
                }
                int count = datasamplelist.Rows.Count;
                this.datasamplelist.Rows.Add(1);
                this.datasamplelist.Rows[count].Cells[0].Value = this.txtsamplebar.Text.Trim().ToUpper();
                txtfactsampleqty.Text = datasamplelist.Rows.Count.ToString();
                txtsamplebar.Text = "";

                if (txtsampleqty.Text == this.txtfactsampleqty.Text.Trim())
                {
                    txtQC.Focus();
                }
                else
                {
                    this.lblinfo.Text = "";
                    txtsamplebar.Enabled = true;
                    txtsamplebar.Focus();
                }
            }
        }
        private void datasamplelist_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            txtfactsampleqty.Text = (int.Parse(txtfactsampleqty.Text) - 1).ToString();
        }
        private string sampleqty(int lotqty, string AQLValue)
        {
            string ssampleqty = @"Select case when Sampleqty>lotfactqty then lotfactqty else Sampleqty end as Sampleqty,Code,AQLValue,AQL,Ac,Re,DirectCode,CheckLevel from IQC_TestSTD105ECode c ";
            ssampleqty += "  inner join ";
            ssampleqty += " (Select " + lotqty + " lotfactqty,AQLValue,AQL,CheckLevel,Ac,Re,DirectCode from IQC_TestSTD105ECheckSet i inner join IQC_TestAQLRcSet s on i.Code=s.Code  where LotSizemin<=" + lotqty.ToString() + " and LotSizemax>=" + lotqty.ToString() + " and CheckLevel='II' and AQLValue='" + AQLValue + "') s on c.Code=s.DirectCode";

            DataSet dssampleqty = Common.DbAccess.SelectBySql(ssampleqty);
            string sampleqty = "0";
            if (dssampleqty != null && dssampleqty.Tables.Count > 0 && dssampleqty.Tables[0].Rows.Count > 0)
                sampleqty = dssampleqty.Tables[0].Rows[0]["Sampleqty"].ToString();
            return sampleqty;
        }
        private void txtsendqty_Leave(object sender, EventArgs e)
        {
            int i = 0;
            if (txtsendqty.Text != "" && int.TryParse(txtsendqty.Text, out i))
            {
                txtAQL.SelectedValue = "0.65";
                txtsampleqty.Text = sampleqty(int.Parse(txtsendqty.Text), txtAQL.SelectedValue.ToString());
                DataTable dt = bindtypes();
                DataRow[] rw = dt.Select("Definetype='无需扫描条码编码'");
                if (rw.Length > 0)
                {
                    txtQC.Focus();
                }
                else
                {
                    txtsamplebar.Focus();
                    ifscan = "是";
                }
            }
            else
            {
                return;
            }
        }
        private void txtAQL_SelectedIndexChanged(object sender, EventArgs e)
        {
            int i = 0;
            if (txtAQL.SelectedValue == null) return;
            if (!int.TryParse(txtsendqty.Text, out i)) return;
            txtsampleqty.Text = sampleqty(int.Parse(txtsendqty.Text), txtAQL.SelectedValue.ToString());
            txtsamplebar.Focus();
        }
        private void cbok_CheckedChanged(object sender, EventArgs e)
        {
            if (cbok.Checked)
            {
                cbok.Text = "OK";
                this.txtremarks.Enabled = true;
                txtremarks.Text = "";
                cbNG.Checked = false;
                txtNGqty.Enabled = false;
            }
        }
        private void cbNG_CheckedChanged(object sender, EventArgs e)
        {
            if (cbNG.Checked)
            {
                cbNG.Text = "NG";
                this.txtremarks.Enabled = true;
                cbok.Checked = false;
                this.txtNGqty.Enabled = true;
                txtNGqty.SelectAll();
            }
        }
        private void btnsearchbydate_Click(object sender, EventArgs e)
        {
            databind.DataSource = null;
            string sbegindate = "", senddate = "";
            if (txtbegindate.Checked)
                sbegindate = txtbegindate.Value.ToString();
            if (txtenddate.Checked)
                senddate = txtenddate.Value.ToString();

            databind.DataSource = ic.AddOQCTestRecordSearch("按日期查询", txtproductcode.Text, txttestitem.SelectedValue == null ? "" : txttestitem.SelectedValue.ToString(), txtcuscode.Text,
                    txtcustomer.SelectedValue == null ? "" : txtcustomer.SelectedValue.ToString(), txtorg_id.SelectedValue.ToString(), txtworkno.Text, txtline.SelectedValue == null ? "" : txtline.SelectedValue.ToString(),
                    txttesttools.SelectedValue == null ? "" : txttesttools.SelectedValue.ToString(), txtsendqty.Text, txtsampleqty.Text, txtAQL.SelectedValue.ToString(), txtAQL.Text, "", txtremarks.Text, Login.username,
                    txtQC.Text, txtQE.SelectedValue == null ? "" : txtQE.SelectedValue.ToString(), txtmaster.SelectedValue == null ? "" : txtmaster.SelectedValue.ToString(),
                    txtstate.SelectedValue == null ? "" : txtstate.SelectedValue.ToString(), "", "", sbegindate, senddate).Tables[0];

        }
        public string AddOQCTestRecord(string opertype, string productcode, string testitem, string cuscode, string customer, string org_id, string workno, string line,
      string testtool, string sendqty, string sampleqty, string AQLValue, string AQL, string testresult, string testremarks, string userid, string qc,
      string qe, string master, string states, string item, string latyper, string begindate, string enddate, SqlConnection sqlConn, SqlTransaction sqlTran)
        {

            SqlCommand sqlCmd;
            sqlCmd = new SqlCommand();
            sqlCmd.CommandType = CommandType.StoredProcedure;
            sqlCmd.CommandText = "OQC_AddTestList";
            sqlCmd.Connection = sqlConn;
            sqlCmd.Transaction = sqlTran;
            SqlParameter[] para = new SqlParameter[25];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@productcode", productcode);
            para[2] = new SqlParameter("@testitem", testitem);
            para[3] = new SqlParameter("@cuscode", cuscode);
            para[4] = new SqlParameter("@customer", customer);
            para[5] = new SqlParameter("@org_id", org_id);
            para[6] = new SqlParameter("@workno", workno);
            para[7] = new SqlParameter("@line", line);
            para[8] = new SqlParameter("@testtool", testtool);
            para[9] = new SqlParameter("@sendqty", sendqty);
            para[10] = new SqlParameter("@sampleqty", sampleqty);
            para[11] = new SqlParameter("@AQLValue", AQLValue);
            para[12] = new SqlParameter("@AQL", AQL);
            para[13] = new SqlParameter("@testresult", testresult);
            para[14] = new SqlParameter("@testremarks", testremarks);
            para[15] = new SqlParameter("@userid", userid);
            para[16] = new SqlParameter("@qc", qc);
            para[17] = new SqlParameter("@qe", qe);
            para[18] = new SqlParameter("@master", master);
            para[19] = new SqlParameter("@states", states);
            para[20] = new SqlParameter("@items", item);
            para[21] = new SqlParameter("@latyper", latyper);
            para[22] = new SqlParameter("@begindate", begindate);
            para[23] = new SqlParameter("@enddate", enddate);
            para[24] = new SqlParameter("@msg", SqlDbType.VarChar, 100);
            para[24].Direction = ParameterDirection.Output;
            foreach (SqlParameter pa in para)
            {
                sqlCmd.Parameters.Add(pa);
            }
            sqlCmd.ExecuteNonQuery();
            return sqlCmd.Parameters["@msg"].Value.ToString();
        }
        private int ExecutesqybyList(string sql, SqlConnection sqlConn, SqlTransaction sqlTran)
        {
            int i = 0;
            if (sql != "")
            {
                SqlCommand sqlCmd;
                sqlCmd = new SqlCommand();
                sqlCmd.CommandType = CommandType.Text;
                sqlCmd.CommandText = sql;
                sqlCmd.Connection = sqlConn;
                sqlCmd.Transaction = sqlTran;
                i = sqlCmd.ExecuteNonQuery();
            }
            else
                i = 1;
            return i;
        }
        private void btnsave_Click(object sender, EventArgs e)
        {
            if (txtworkno.Text == "") return;
            if (txtsendqty.Text == "") return;
            if (txtcustomer.SelectedValue == null || txtline.SelectedValue == null || txtQE.SelectedValue == null
                || txtmaster.SelectedValue == null || txtstate.SelectedValue == null || txttestitem.SelectedValue == null
                || txttesttools.SelectedValue == null || txtline.SelectedValue == null)
            {
                MessageBox.Show("请选择一项内容", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            int m = 1, n = 0;
            if (!int.TryParse(this.txtsendqty.Text == "" ? "1" : this.txtsendqty.Text, out m))
            {
                MessageBox.Show("送检数量必须为数字", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                this.txtsendqty.Text = "";
                this.txtsendqty.Focus();
                return;
            }
            else
            {
                if (m <= 0)
                {
                    MessageBox.Show("送检数量必须为大于0", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    this.txtsendqty.Text = "";
                    this.txtsendqty.Focus();
                    return;
                }
            }
            if (!int.TryParse(txtNGqty.Text, out n))
            {
                MessageBox.Show("NG数量大于0", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                txtNGqty.Text = "0";
                return;
            }
            if (!cbok.Checked && !cbNG.Checked)
            {
                MessageBox.Show("请选择一个检验结果", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string states = "NG";
            if (cbok.Checked)
                states = cbok.Text;
            else if (cbNG.Checked)
                states = cbNG.Text;
            SqlConnection conn = new SqlConnection(Common.DbAccess.connSql);
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            SqlTransaction tran = conn.BeginTransaction();
            string sql = "select getdate(),replace(replace(replace(replace(convert(varchar(30),getdate(),121),'-',''),' ',''),':',''),'.','')";
            string checkdate = Common.DbAccess.SelectBySql(sql).Tables[0].Rows[0][0].ToString();
            string sstr = Common.DbAccess.SelectBySql(sql).Tables[0].Rows[0][1].ToString();
            string exesql = "";
            for (int i = 0; i < datasamplelist.Rows.Count; i++)
            {
                exesql += " Insert into OQC_SampleFactList(workno,org_id, serialno,items,checkdate) values ('" + txtworkno.Text + "','" + txtorg_id.SelectedValue.ToString() + "','" + datasamplelist.Rows[i].Cells[0].Value.ToString() + "','" + sstr + "','" + checkdate + "')" + ";";
            }

            string msg = AddOQCTestRecord("新增", txtproductcode.Text, txttestitem.SelectedValue.ToString(), txtcuscode.Text, txtcustomer.SelectedValue.ToString()
                                    , txtorg_id.SelectedValue.ToString(), txtworkno.Text, txtline.SelectedValue.ToString(), txttesttools.SelectedValue.ToString(), txtsendqty.Text
                                    , txtsampleqty.Text, txtAQL.SelectedValue.ToString(), txtAQL.Text, states, txtremarks.Text, Login.username,
                                    txtQC.Text, txtQE.SelectedValue.ToString(), txtmaster.SelectedValue.ToString(), txtstate.SelectedValue.ToString(), txtNGqty.Text, txtlatyper.Text, sstr, "", conn, tran);
            if (msg.IndexOf("OK") >= 0)
            {
                int k = ExecutesqybyList(exesql, conn, tran);
                if (k > 0)
                {
                    tran.Commit();
                    conn.Close();
                    datasamplelist.Rows.Clear();
                    txtfactsampleqty.Text = "0";
                    txtsendqty.Text = "";
                    txtsampleqty.Text = "";
                    txtQC.Text = "";
                    txtlatyper.Text = "";
                    txtremarks.Text = "";
                    txtworkno.Text = "";
                    ifscan = "否";
                    MessageBox.Show(msg, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    tran.Rollback();
                    conn.Close();
                }

            }
            else
            {
                MessageBox.Show(msg, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            databind.DataSource = ic.AddOQCTestRecordSearch("查询", txtproductcode.Text, txttestitem.SelectedValue == null ? "" : txttestitem.SelectedValue.ToString(), txtcuscode.Text,
                    txtcustomer.SelectedValue == null ? "" : txtcustomer.SelectedValue.ToString(), txtorg_id.SelectedValue.ToString(), txtworkno.Text, txtline.SelectedValue == null ? "" : txtline.SelectedValue.ToString(),
                    txttesttools.SelectedValue == null ? "" : txttesttools.SelectedValue.ToString(), txtsendqty.Text, txtsampleqty.Text, txtAQL.SelectedValue.ToString(), txtAQL.Text, states, txtremarks.Text, Login.username,
                    txtQC.Text, txtQE.SelectedValue == null ? "" : txtQE.SelectedValue.ToString(), txtmaster.SelectedValue == null ? "" : txtmaster.SelectedValue.ToString(), txtstate.SelectedValue == null ? "" :
                    txtstate.SelectedValue.ToString(), "", "", "", "").Tables[0];

        }
        private void btndel_Click(object sender, EventArgs e)
        {
            if (databind.SelectedRows.Count <= 0) return;
            for (int k = databind.SelectedRows.Count; k > 0; k--)
            {
                string msg = ic.AddOQCTestRecord("删除", databind.SelectedRows[k - 1].Cells["Hytera编码"].Value.ToString(), databind.SelectedRows[k - 1].Cells["测试项目"].Value.ToString(), databind.SelectedRows[k - 1].Cells["客户型号"].Value.ToString(), "", databind.SelectedRows[k - 1].Cells["组织"].Value.ToString(),
                                                databind.SelectedRows[k - 1].Cells["工单号"].Value.ToString(), "", "", "", "", "0", "0", "0", "", "0", "0", "", "0", "0", databind.SelectedRows[k - 1].Cells["序号"].Value.ToString(), "", "", "");
            }
            databind.DataSource = ic.AddOQCTestRecordSearch("查询", txtproductcode.Text, txttestitem.SelectedValue == null ? "" : txttestitem.SelectedValue.ToString(), txtcuscode.Text,
                    txtcustomer.SelectedValue == null ? "" : txtcustomer.SelectedValue.ToString(), txtorg_id.SelectedValue.ToString(), txtworkno.Text, txtline.SelectedValue == null ? "" : txtline.SelectedValue.ToString(),
                    txttesttools.SelectedValue == null ? "" : txttesttools.SelectedValue.ToString(), txtsendqty.Text, txtsampleqty.Text, txtAQL.SelectedValue.ToString(), txtAQL.Text, "", txtremarks.Text, Login.username,
                    txtQC.Text, txtQE.SelectedValue == null ? "" : txtQE.SelectedValue.ToString(), txtmaster.SelectedValue == null ? "" : txtmaster.SelectedValue.ToString(),
                    txtstate.SelectedValue == null ? "" : txtstate.SelectedValue.ToString(), "", "", "", "").Tables[0];
        }
        private void btnsearch_Click(object sender, EventArgs e)
        {
            databind.DataSource = null;
            databind.DataSource = ic.AddOQCTestRecordSearch("查询", txtproductcode.Text, txttestitem.SelectedValue == null ? "" : txttestitem.SelectedValue.ToString(), txtcuscode.Text,
                    txtcustomer.SelectedValue == null ? "" : txtcustomer.SelectedValue.ToString(), txtorg_id.SelectedValue.ToString(), txtworkno.Text, txtline.SelectedValue == null ? "" : txtline.SelectedValue.ToString(),
                    txttesttools.SelectedValue == null ? "" : txttesttools.SelectedValue.ToString(), txtsendqty.Text, txtsampleqty.Text, txtAQL.SelectedValue.ToString(), txtAQL.Text, "", txtremarks.Text, Login.username,
                    txtQC.Text, txtQE.SelectedValue == null ? "" : txtQE.SelectedValue.ToString(), txtmaster.SelectedValue == null ? "" : txtmaster.SelectedValue.ToString(),
                    txtstate.SelectedValue == null ? "" : txtstate.SelectedValue.ToString(), "", "", "", "").Tables[0];

        }
        public void DataToExcel(DataGridView m_DataView)
        {
            SaveFileDialog kk = new SaveFileDialog();
            kk.Title = "保存EXECL文件";
            kk.Filter = "EXECL文件(*.xls) |*.xls";
            kk.FilterIndex = 1;

            if (kk.ShowDialog() == DialogResult.OK)
            {
                string FileName = kk.FileName;
                if (File.Exists(FileName))
                    File.Delete(FileName);
                FileStream objFileStream;
                StreamWriter objStreamWriter;
                string strLine = "";
                objFileStream = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.Write);
                objStreamWriter = new StreamWriter(objFileStream, System.Text.Encoding.Unicode);

                for (int i = 0; i < m_DataView.Columns.Count; i++)
                {
                    if (m_DataView.Columns[i].Visible == true)
                    {
                        strLine = strLine + m_DataView.Columns[i].HeaderText.ToString() + Convert.ToChar(9);
                    }
                }
                objStreamWriter.WriteLine(strLine);
                strLine = "";

                for (int i = 0; i < m_DataView.Rows.Count; i++)
                {
                    if (m_DataView.Columns[0].Visible == true)
                    {
                        if (m_DataView.Rows[i].Cells[0].Value == null)
                            strLine = strLine + " " + Convert.ToChar(9);
                        else
                            strLine = strLine + m_DataView.Rows[i].Cells[0].Value.ToString() + Convert.ToChar(9);
                    }
                    for (int j = 1; j < m_DataView.Columns.Count; j++)
                    {

                        if (m_DataView.Columns[j].Visible == true)
                        {
                            if (m_DataView.Rows[i].Cells[j].Value == null)
                                strLine = strLine + " " + Convert.ToChar(9);
                            else
                            {
                                string rowstr = "";
                                rowstr = m_DataView.Rows[i].Cells[j].Value.ToString();
                                if (rowstr.IndexOf("\r\n") > 0)
                                    rowstr = rowstr.Replace("\r\n", " ");
                                if (rowstr.IndexOf("\t") > 0)
                                    rowstr = rowstr.Replace("\t", " ");
                                if (rowstr.IndexOf("\n") > 0)
                                    rowstr = rowstr.Replace("\n", " ");

                                strLine = strLine + rowstr + Convert.ToChar(9);
                            }
                        }
                    }
                    objStreamWriter.WriteLine(strLine);
                    strLine = "";
                }
                objStreamWriter.Close();
                objFileStream.Close();

            }
        }
        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (databind.Rows.Count <= 0) return;
             DataToExcel(databind);
        }
        private void databind_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (databind.CurrentRow.Cells["实抽数量"].Value.ToString() != "")
                {
                    string item = databind.CurrentRow.Cells["序号"].Value.ToString();
                    OQCSampleList OL = new OQCSampleList(item);
                    DialogResult dr = OL.ShowDialog();
                }
            }
            catch
            {
            }
        }
        private void databind_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            DataGridViewRow dgr = databind.Rows[e.RowIndex];
            try
            {
                if (dgr.Cells["检验结果"].Value.ToString() == "NG")
                {
                    dgr.DefaultCellStyle.BackColor = Color.Yellow;
                }
            }
            catch
            {

            }
        }
        private void txtsn_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string sql = "select lotno from Pack_Maintain where serialno='" + txtsn.Text.Trim() + "'";
                DataTable dt = Common.DbAccess.SelectBySql(sql).Tables[0];
                if (dt.Rows.Count > 0)
                {
                    txtcolorsn.Text = "";
                    txtcolorsn.Focus();
                    lotno = dt.Rows[0]["lotno"].ToString();
                    txtsn.Enabled = false;
                }
                else
                {
                    lblinfo_2.Text = "机身号:" + txtsn.Text + ",不存在!";
                    lblinfo_2.ForeColor = Color.Red;
                    txtsn.Text = "";
                    txtsn.Focus();
                    txtsn.Enabled = true;
                    lotno = "";
                }
            }
        }
        private void txtcolorsn_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txtcolorsn.Text != txtsn.Text)
                {
                    lblinfo_2.Text = txtcolorsn.Text + "彩合号与机身号不一致!";
                    lblinfo_2.ForeColor = Color.Red;
                    txtcolorsn.Text = "";
                    txtcolorsn.Focus();
                    txtcolorsn.Enabled = true;
                }
                else
                {
                    txtboxsn.Text = "";
                    txtboxsn.Focus();
                    txtcolorsn.Enabled = false;
                }
            }
        }
        private void txtboxsn_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txtboxsn.Text.Trim().ToUpper() != lotno.ToUpper())
                {
                    lblinfo_2.Text = txtboxsn.Text + "外箱号不正确,正确为:" + lotno;
                    lblinfo_2.ForeColor = Color.Red;
                    txtboxsn.Enabled = true;
                    txtboxsn.Text = "";
                    txtboxsn.Focus();
                }
                else
                {
                    string sql = "if not exists(select 1 from IQC_QACheck where SN='" + txtsn.Text + "'" + ")";
                    sql += " insert into IQC_QACheck(SN,CorSN,LotNo,operuser,operdate) values('" + txtsn.Text + "','" + txtcolorsn.Text + "','" + txtboxsn.Text + "','" + Login.userId + "',getdate()" + ")";
                    bool F = Common.DbAccess.ExecuteSql(sql);
                    if (F)
                    {
                        databind.Rows.Add(1);
                        int count = databind.Rows.Count;
                        databind.Rows[count - 1].Cells["SN"].Value = txtsn.Text;
                        databind.Rows[count - 1].Cells["ColorSN"].Value = txtcolorsn.Text;
                        databind.Rows[count - 1].Cells["BoxSN"].Value = txtboxsn.Text;
                        lblinfo_2.Text = txtsn.Text + "核对成功";
                        lblinfo_2.ForeColor = Color.Blue;
                        txtboxsn.Text = "";
                        txtcolorsn.Text = "";
                        txtcolorsn.Enabled = true;
                        txtsn.Text = "";
                        txtsn.Enabled = true;
                        txtsn.Focus();
                    }
                    else
                    {
                        for (int k = 0; k < databind.Rows.Count; k++)
                        {
                            if (databind.Rows[k].Cells["SN"].Value.ToString() == txtsn.Text.ToUpper().Trim())
                            {
                                lblinfo_2.Text = txtsn.Text + "已经核对过";
                                lblinfo_2.ForeColor = Color.Blue;
                                txtboxsn.Text = "";
                                txtcolorsn.Text = "";
                                txtcolorsn.Enabled = true;
                                txtsn.Text = "";
                                txtsn.Enabled = true;
                                txtsn.Focus();
                                return;
                            }

                        }
                        databind.Rows.Add(1);
                        int count = databind.Rows.Count;
                        databind.Rows[count - 1].Cells["SN"].Value = txtsn.Text;
                        databind.Rows[count - 1].Cells["ColorSN"].Value = txtcolorsn.Text;
                        databind.Rows[count - 1].Cells["BoxSN"].Value = txtboxsn.Text;
                        lblinfo_2.Text = txtsn.Text + "已经核对过";
                        lblinfo_2.ForeColor = Color.Blue;
                        txtboxsn.Text = "";
                        txtcolorsn.Text = "";
                        txtcolorsn.Enabled = true;
                        txtsn.Text = "";
                        txtsn.Enabled = true;
                        txtsn.Focus();
                    }

                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (txtsn.Text == "" && txtcolorsn.Text == "" && txtboxsn.Text == "" && txtworkno_1.Text == "") return;
            databind.Rows.Clear();
            string sql = "select SN,CorSN,q.LotNo,workno from IQC_QACheck q inner join Pack_Maintain p on q.SN=p.serialno where 1=1";
            if (txtsn.Text != "")
                sql += " and SN='" + this.txtsn.Text + "'";
            if (txtcolorsn.Text != "")
                sql += " and CorSN='" + txtcolorsn.Text + "'";
            if (txtboxsn.Text != "")
                sql += " and q.LotNo='" + txtboxsn.Text + "'";
            if (this.txtworkno_1.Text != "")
                sql += " and workno='" + txtworkno_1.Text + "'";
            DataTable dt = Common.DbAccess.SelectBySql(sql).Tables[0];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                databind.Rows.Add(1);
                int count = databind.Rows.Count;
                databind.Rows[count - 1].Cells["SN"].Value = dt.Rows[i]["SN"].ToString();
                databind.Rows[count - 1].Cells["ColorSN"].Value = dt.Rows[i]["CorSN"].ToString();
                databind.Rows[count - 1].Cells["BoxSN"].Value = dt.Rows[i]["LotNo"].ToString();
            }
        }
    }
}