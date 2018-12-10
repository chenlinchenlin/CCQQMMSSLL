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
using System.IO;
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraEditors;
using System.Data.SqlClient;

namespace DX_QMS
{
    public partial class ESDTestNew : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private Microsoft.Office.Interop.Excel.ApplicationClass appClsExcel = null;
        DataTable dtsubitem = null;
        IQC ic = new IQC();
        public ESDTestNew()
        {
            InitializeComponent();
            txtstates.SelectedIndex = 0;
            DataTable dt, dt1, dt2, dt3;
            try
            {
                dt = bindTypesByName("楼层");
                txtblock.Properties.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    txtblock.Properties.Items.Add(row["Dvalue"]);
                }
               ////// txtblock.SelectedIndex = 0;
            }
            catch
            {
            }
            setRule();
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }
        private DataTable bindTypesByName(string types)
        {
            string sql = "select Dvalue from ESD_TypeDefine where Dtype='" + types + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        private void setRule()
        {
            string post = "";
            if (Login.manager != null && Login.manager != "")
            {
                post = Login.manager;
            }
            else
            {
                post = Login.post;
            }
            Dictionary<string, bool> dic = GroupPermission.QMS_SelectRulesForForm(post, "测试记录");
            this.btnsave.Enabled = bool.Parse(dic["hasInsert"].ToString());
            this.btndel.Enabled = bool.Parse(dic["hasDelete"].ToString());
            this.btntoexcel.Enabled = bool.Parse(dic["hasPrint"].ToString());
        }


        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                if (ActiveControl is TextEdit || ActiveControl is ComboBoxEdit || ActiveControl is DateEdit)
                {
                    SendKeys.Send("{tab}");
                    return true;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        private void ESDTestNew_FormClosing(object sender, FormClosingEventArgs e)
        {

        }
        private void bindSubType(string block, string line, string sType)
        {
            string sql = "select DsubType,UpValue, LowValue from ESD_TestProgSet where line='" + line + "' and block='" + block + "' and Dtype='" + sType + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt.Rows.Count > 0)
            {
                dtsubitem = dt;
                txttestsubitem.Properties.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    txttestsubitem.Properties.Items.Add(row["DsubType"]);
                }
                txttestsubitem.SelectedIndex = 0;
                if (txttestsubitem.Properties.Items.Count == 1)
                {
                    txttestsubitem.SelectedIndex = 0;
                    DataRow[] rw = dtsubitem.Select("DsubType='" + this.txttestsubitem.Text + "'");
                    txtUppervalue.Text = rw[0]["UpValue"].ToString();
                    txtLowervalue.Text = rw[0]["LowValue"].ToString();
                }
                else
                {
                    txttestsubitem.Focus();
                    return;
                }
            }
            else
            {
                txttestsubitem.Properties.Items.Clear();
            }
        }
        private void bindType(string block, string line)
        {
            string sql = "select distinct  Dtype from ESD_TestProgSet where line='" + line + "' and block='" + block + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt.Rows.Count > 0)
            {
                DataTable newdt = dt.Clone();
                newdt.Rows.Add(dt.Rows[0].ItemArray);
                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    bool flag = true;
                    foreach (DataRow dr in newdt.Rows)
                    {
                        if (dt.Rows[i]["Dtype"].ToString() == dr["Dtype"].ToString())
                        {
                            flag = false;
                            continue;
                        }
                    }
                    if (flag)
                        newdt.Rows.Add(dt.Rows[i].ItemArray);
                }
                txttesttype.Properties.Items.Clear();
                foreach (DataRow row in newdt.Rows)
                {
                    txttesttype.Properties.Items.Add(row["Dtype"]);
                }
                txttesttype.SelectedIndex = 0;

                bindSubType(txtblock.Text, txtline.Text, txttesttype.Text);
            }
            else
            {
                txttesttype.Properties.Items.Clear();
                txttestsubitem.Properties.Items.Clear();
            }
        }
        private void bindLine(string block)
        {
            string sql = "select distinct line from ESD_TestProgSet where Block='" + block + "'";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt.Rows.Count > 0)
            {
                DataTable newdt = dt.Clone();
                newdt.Rows.Add(dt.Rows[0].ItemArray);
                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    bool flag = true;
                    foreach (DataRow dr in newdt.Rows)
                    {
                        if (dt.Rows[i]["line"].ToString() == dr["line"].ToString())
                        {
                            flag = false;
                            continue;
                        }
                    }
                    if (flag)
                        newdt.Rows.Add(dt.Rows[i].ItemArray);
                }
                txtline.Properties.Items.Clear();
                foreach (DataRow row in newdt.Rows)
                {
                    txtline.Properties.Items.Add(row["line"]);
                }
                txtline.SelectedIndex = 0;

                bindType(txtblock.Text, txtline.Text);
            }
            else
            {
                txtline.Properties.Items.Clear();
                txttesttype.Properties.Items.Clear();
                txttestsubitem.Properties.Items.Clear();
            }
        }
        private void txtblock_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtblock.Text == "")
                return;
            bindLine(txtblock.Text);
        }
        private void txtline_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtline.Text == "")
                return;
            bindType(txtblock.Text, txtline.Text);
        }
        private void txttesttype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txttesttype.Text == null)
                return;
            bindSubType(txtblock.Text, txtline.Text, txttesttype.Text);
        }
        private void txttestsubitem_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txttestsubitem.Text  == "")
                return;
            try
            {
                DataRow[] rw = dtsubitem.Select("DsubType='" + this.txttestsubitem.Text + "'");
                txtUppervalue.Text = rw[0]["UpValue"].ToString();
                txtLowervalue.Text = rw[0]["LowValue"].ToString();
            }
            catch
            {
            }
        }
        public abstract class ScienceCount
        {
            public static string KXJSF(double num)
            {
                double bef = System.Math.Abs(num);
                int aft = 0;
                while (bef >= 10 || (bef < 1 && bef != 0))
                {
                    if (bef >= 10)
                    {
                        bef = bef / 10;
                        aft++;
                    }
                    else
                    {
                        bef = bef * 10;
                        aft--;
                    }
                }
                return string.Concat(num >= 0 ? "" : "-", ReturnBef(bef), "E", ReturnAft(aft));
            }
            public static string ReturnBef(double bef)
            {
                if (bef.ToString() != null)
                {
                    char[] arr = bef.ToString().ToCharArray();
                    switch (arr.Length)
                    {
                        case 1:
                        case 2: return string.Concat(arr[0], ".", "00"); break;
                        case 3: return string.Concat(arr[0] + "." + arr[2] + "0"); break;
                        default: return string.Concat(arr[0] + "." + arr[2] + arr[3]); break;
                    }
                }
                else
                {
                    return "000";
                }
            }
            public static string ReturnAft(int aft)
            {
                if (aft.ToString() != null)
                {
                    string end;
                    char[] arr = System.Math.Abs(aft).ToString().ToCharArray();
                    switch (arr.Length)
                    {
                        case 1: end = "00" + arr[0]; break;
                        case 2: end = "0" + arr[0] + arr[1]; break;
                        default: end = System.Math.Abs(aft).ToString(); break;
                    }
                    return string.Concat(aft >= 0 ? "+" : "-", end);
                }
                else
                {
                    return "+000";
                }
            }
        }
        string testresult = "";
        private void txttestvalue_Leave(object sender, EventArgs e)
        {
            if (!(txttesttype.Text.Contains("离子") || txttesttype.Text.Contains("静电监控")))
            {
                float m = 0;
                if (txttestvalue.Text == "")
                    return;
                if (!float.TryParse(txttestvalue.Text, out m)) return;

                this.txttestvalueSI.Text = ScienceCount.KXJSF(double.Parse(this.txttestvalue.Text));

                if (m < float.Parse(txtLowervalue.Text) || m > float.Parse(txtUppervalue.Text))
                {
                    cbNG.Checked = true;
                    testresult = cbNG.Text;
                    cbOK.Checked = false;
                }
                if (m >= float.Parse(txtLowervalue.Text) && m <= float.Parse(txtUppervalue.Text))
                {
                    cbOK.Checked = true;
                    testresult = cbOK.Text;
                    cbNG.Checked = false;
                }
            }
        }

        public string AddNewTestESDnew(string opertype, string block, string line, string testdate, string testtype, string testsubtype, string testvalue, string userid, int item, string finaltestdate, string testresult, string testremarks,string positivetesttime,string negativetesttime,string balancedvoltages)
        {
            SqlParameter[] para = new SqlParameter[16];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@block", block);
            para[2] = new SqlParameter("@line", line);
            para[3] = new SqlParameter("@testdate", testdate);
            para[4] = new SqlParameter("@testtype", testtype);
            para[5] = new SqlParameter("@testsubtype", testsubtype);
            para[6] = new SqlParameter("@testvalue", testvalue);
            para[7] = new SqlParameter("@userid", userid);
            para[8] = new SqlParameter("@item", item);
            para[9] = new SqlParameter("@finaltestdate", finaltestdate);
            para[10] = new SqlParameter("@testresult", testresult);
            para[11] = new SqlParameter("@testremarks", testremarks);
            para[12] = new SqlParameter("@positivetesttime", positivetesttime);
            para[13] = new SqlParameter("@negativetesttime", negativetesttime);
            para[14] = new SqlParameter("@balancedvoltages", balancedvoltages);
            para[15] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[15].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "ESD_InsertTestESDnew", para);
            return para[15].Value.ToString();
        }



        private void btnsave_Click(object sender, EventArgs e)
        {
            if (txtblock.Text == "" || txtline.Text == "" || txttesttype.Text == "" || txttestsubitem.Text == "")
                return;
            string msg = "",result = "";

            if (txttestsubitem.Text.Contains("离子"))
            {
                float m = 0;
                if (!float.TryParse(txtpositivetesttime.Text == "" ? "0" : txtpositivetesttime.Text, out m))
                {
                    MessageBox.Show("正消散时间值不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txtpositivetesttime.Text = "";
                    txtpositivetesttime.Focus();
                    return;
                }
                if (!float.TryParse(txtnegativetesttime.Text == "" ? "0" : txtnegativetesttime.Text, out m))
                {
                    MessageBox.Show("负消散时间值不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txtnegativetesttime.Text = "";
                    txtnegativetesttime.Focus();
                    return;
                }
                if (!float.TryParse(txtbalancedvoltages.Text == "" ? "0" : txtbalancedvoltages.Text, out m))
                {
                    MessageBox.Show("平衡电压值不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txtbalancedvoltages.Text = "";
                    txtbalancedvoltages.Focus();
                    return;
                }
                if (cbOK.Checked == false && cbNG.Checked == false)
                {
                    MessageBox.Show("请选择一个测试结果", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (cbOK.Checked == true && cbNG.Checked == false)
                    result = "OK";
                if (cbOK.Checked == false && cbNG.Checked == true)
                    result = "NG";
                if (result == "")
                    return;

                msg = AddNewTestESDnew("新增离子化设备", txtblock.Text, txtline.Text, datetime.DateTime.ToString(), txttesttype.Text, txttestsubitem.Text,
                       txttestvalue.Text == "" ? "0" : txttestvalue.Text, Login.username, 0, datetime.DateTime.ToString(),result, txtremark.Text, txtpositivetesttime.Text + " S", txtnegativetesttime.Text + " S", txtbalancedvoltages.Text + " V");
            }
            else if (txttestsubitem.Text.Contains("静电监控"))
            {
                if (cbOK.Checked == false && cbNG.Checked ==false )
                {
                    MessageBox.Show("请选择一个测试结果", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (cbOK.Checked == true && cbNG.Checked == false)                
                    result = "OK";
                if (cbOK.Checked == false && cbNG.Checked == true)
                    result = "NG";
                if (result == "")
                    return;
                msg = ic.AddNewTestESD("新增", txtblock.Text, txtline.Text, datetime.DateTime.ToString(), txttesttype.Text, txttestsubitem.Text,
                    txttestvalue.Text == "" ? "0" : txttestvalue.Text, Login.username, 0, datetime.DateTime.ToString(),result, txtremark.Text);

            }
            else
            {
                if (txttestvalue.Text == "")
                {
                    MessageBox.Show("请输入测试值","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    return;
                }
                float n = 0;
                if (!float.TryParse(txtLowervalue.Text == "" ? "0" : txtLowervalue.Text, out n))
                {
                    MessageBox.Show("测试值范围值不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txtLowervalue.Text = "";
                    txtLowervalue.Focus();
                    return;
                }
                if (!float.TryParse(txtUppervalue.Text == "" ? "0" : txtUppervalue.Text, out n))
                {
                    MessageBox.Show("测试值范围值不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txtUppervalue.Text = "";
                    txtUppervalue.Focus();
                    return;
                }
                msg = ic.AddNewTestESD("新增", txtblock.Text, txtline.Text, datetime.DateTime.ToString(), txttesttype.Text, txttestsubitem.Text,
                      txttestvalue.Text == "" ? "0" : txttestvalue.Text, Login.username, 0, datetime.DateTime.ToString(), testresult, txtremark.Text);

            }
           if (msg.IndexOf("OK") >= 0)
           {
                lblinfo.Text = msg;
                lblinfo.ForeColor = Color.Blue;
           }
           else
           {
                MessageBox.Show(msg, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
           }
            btnsearch_Click(sender, e);
            txtLowervalue.Text = "";
            txtUppervalue.Text = "";
        }

        private void btndel_Click(object sender, EventArgs e)
        {
            if (gridView.FocusedRowHandle < 0)
                return;
            if (gridView.GetSelectedRows().Length < 1)
                return;
            for (int k = gridView.GetSelectedRows().Length; k > 0; k--)
            {

                DataRow dr = gridView.GetDataRow(gridView.GetSelectedRows()[k - 1]);

                string msg = ic.AddNewTestESD("删除", dr["楼层"].ToString(), dr["线别"].ToString(),
                         "", dr["测试类别"].ToString(), dr["测试子项目"].ToString(), "", "", 0, dr["测试日期"].ToString(), "", "");
            }
            btnsearch_Click(sender, e);
        }
        private DataTable BindESDRecord(string sblock, string sline, string testtype, string testsubitem, string barcode)
        {
            string sql = "select f.lotno 条码,l.Block 楼层,l.line 线别, Dtype 测试类别, DsubType 测试子项目,s.UpValue 上限, s.LowValue 下限,checkcycle 测试周期, ";
            sql += " case when NextFactTdate is null then DATEDIFF(day,getdate(),(isnull(s.checkcycle,90)+FinalTestDate)) else 90 end 距离天数,";
            sql += " Testvalue 测试值, FinalTestDate 测试日期,dateadd(day,checkcycle,FinalTestDate) 下次测试日期,NextFactTdate 实际测试日期,TestResult 测试结果, TestRemarks 备注,TestOperUser 测试人,s.leadtime 提前期 ";
            sql += " ,case when NextFactTdate is null and DATEDIFF(day,getdate(),(isnull(s.checkcycle,90)+FinalTestDate))<0 then '已过期' when NextFactTdate is null and DATEDIFF(day,getdate(),(isnull(s.checkcycle,90)+FinalTestDate))>0 and DATEDIFF(day,getdate(),(isnull(s.checkcycle,90)+FinalTestDate))<s.leadtime ";
            sql += " then '快过期' when NextFactTdate is not null and DATEDIFF(day,NextFactTdate,(isnull(s.checkcycle,90)+FinalTestDate))>0 then '正常' ";
            sql += " when NextFactTdate is not null and DATEDIFF(day,NextFactTdate,(isnull(s.checkcycle,90)+FinalTestDate))<0 then '已过期' else '正常' end as 状态";
            sql += " from ESD_TestList l inner join ESD_LotInfo f on l.block=f.block and l.line=f.line and l.testtype=f.testtype and l.testsubitem=f.testsubitem left join ESD_TestProgSet s on l.block=s.block and l.line=s.line and l.testtype=s.Dtype and l.testsubitem=s.DsubType ";
            sql += " where l.Block= case '" + sblock + "' when '' then l.Block else '" + sblock + "' end";
            sql += " and l.line= case '" + sline + "' when '' then l.line else '" + sline + "' end";
            sql += " and Dtype= case '" + testtype + "' when '' then Dtype else '" + testtype + "' end";
            sql += " and DsubType= case '" + testsubitem + "' when '' then DsubType else '" + testsubitem + "' end";
            sql += " and f.lotno= case '" + barcode + "' when '' then f.lotno else '" + barcode + "' end  order by l.TestOperDate desc  ";
            return DbAccess.SelectBySql(sql).Tables[0];
        }

        private DataTable selectliRecord(string sblock, string sline, string testtype, string testsubitem, string barcode)
        {
            string sql = "select f.lotno 条码,l.Block 楼层,l.line 线别, Dtype 测试类别, DsubType 测试子项目,'+35' 平衡电压上限, '-35' 平衡电压下限,'≤5S' 正负消散时间要求,l.positivetesttime 正消散时间,l.negativetesttime 负消散时间,l.balancedvoltages 平衡电压,checkcycle 测试周期, ";
            sql += " case when NextFactTdate is null then DATEDIFF(day,getdate(),(isnull(s.checkcycle,90)+FinalTestDate)) else 90 end 距离天数,";
            sql += " FinalTestDate 测试日期,dateadd(day,checkcycle,FinalTestDate) 下次测试日期,NextFactTdate 实际测试日期,TestResult 测试结果, TestRemarks 备注,TestOperUser 测试人,s.leadtime 提前期, ";
            sql += " case when NextFactTdate is null and DATEDIFF(day,getdate(),(isnull(s.checkcycle,90)+FinalTestDate))<0 then '已过期' when NextFactTdate is null and DATEDIFF(day,getdate(),(isnull(s.checkcycle,90)+FinalTestDate))>0 and DATEDIFF(day,getdate(),(isnull(s.checkcycle,90)+FinalTestDate))<s.leadtime ";
            sql += " then '快过期' when NextFactTdate is not null and DATEDIFF(day,NextFactTdate,(isnull(s.checkcycle,90)+FinalTestDate))>0 then '正常' ";
            sql += " when NextFactTdate is not null and DATEDIFF(day,NextFactTdate,(isnull(s.checkcycle,90)+FinalTestDate))<0 then '已过期' else '正常' end as 状态";
            sql += " from ESD_TestList l inner join ESD_LotInfo f on l.block=f.block and l.line=f.line and l.testtype=f.testtype and l.testsubitem=f.testsubitem left join ESD_TestProgSet s on l.block=s.block and l.line=s.line and l.testtype=s.Dtype and l.testsubitem=s.DsubType ";
            sql += " where l.Block= case '" + sblock + "' when '' then l.Block else '" + sblock + "' end";
            sql += " and l.line= case '" + sline + "' when '' then l.line else '" + sline + "' end";
            sql += " and Dtype= case '" + testtype + "' when '' then Dtype else '" + testtype + "' end";
            sql += " and DsubType= case '" + testsubitem + "' when '' then DsubType else '" + testsubitem + "' end";
            sql += " and f.lotno= case '" + barcode + "' when '' then f.lotno else '" + barcode + "' end  order by l.TestOperDate desc  ";
            return DbAccess.SelectBySql(sql).Tables[0];
        }

        private DataTable selectjingdianRecord(string sblock, string sline, string testtype, string testsubitem, string barcode)
        {
            string sql = "select f.lotno 条码,l.Block 楼层,l.line 线别, Dtype 测试类别, DsubType 测试子项目,s.UpValue 静电上限, s.LowValue 静电下限,checkcycle 测试周期, ";
            sql += " case when NextFactTdate is null then DATEDIFF(day,getdate(),(isnull(s.checkcycle,90)+FinalTestDate)) else 90 end 距离天数,";
            sql += " FinalTestDate 测试日期,dateadd(day,checkcycle,FinalTestDate) 下次测试日期,NextFactTdate 实际测试日期,TestResult 测试结果, TestRemarks 备注,TestOperUser 测试人,s.leadtime 提前期, ";
            sql += " case when NextFactTdate is null and DATEDIFF(day,getdate(),(isnull(s.checkcycle,90)+FinalTestDate))<0 then '已过期' when NextFactTdate is null and DATEDIFF(day,getdate(),(isnull(s.checkcycle,90)+FinalTestDate))>0 and DATEDIFF(day,getdate(),(isnull(s.checkcycle,90)+FinalTestDate))<s.leadtime ";
            sql += " then '快过期' when NextFactTdate is not null and DATEDIFF(day,NextFactTdate,(isnull(s.checkcycle,90)+FinalTestDate))>0 then '正常' ";
            sql += " when NextFactTdate is not null and DATEDIFF(day,NextFactTdate,(isnull(s.checkcycle,90)+FinalTestDate))<0 then '已过期' else '正常' end as 状态";
            sql += " from ESD_TestList l inner join ESD_LotInfo f on l.block=f.block and l.line=f.line and l.testtype=f.testtype and l.testsubitem=f.testsubitem left join ESD_TestProgSet s on l.block=s.block and l.line=s.line and l.testtype=s.Dtype and l.testsubitem=s.DsubType ";
            sql += " where l.Block= case '" + sblock + "' when '' then l.Block else '" + sblock + "' end";
            sql += " and l.line= case '" + sline + "' when '' then l.line else '" + sline + "' end";
            sql += " and Dtype= case '" + testtype + "' when '' then Dtype else '" + testtype + "' end";
            sql += " and DsubType= case '" + testsubitem + "' when '' then DsubType else '" + testsubitem + "' end";
            sql += " and f.lotno= case '" + barcode + "' when '' then f.lotno else '" + barcode + "' end  order by l.TestOperDate desc  ";
            return DbAccess.SelectBySql(sql).Tables[0];
        }



        private void btnsearch_Click(object sender, EventArgs e)
        {       
            string sblock, sline, stype, subtype;
            if (txtblock.Text == "")
                sblock = "";
            else
                sblock = txtblock.Text;
            if (txtline.Text == "")
                sline = "";
            else
                sline = txtline.Text;
            if (txttesttype.Text == "")
                stype = "";
            else
                stype = txttesttype.Text;
            subtype = "";
            gridView.Columns.Clear();
            databind.DataSource = null;
            DataTable dt = null;
            if (txttesttype.Text.Contains("静电监控"))
            {
                dt = selectjingdianRecord(sblock, sline, stype, subtype, txtbarcode.Text);
            }
            else if (txttesttype.Text.Contains("离子"))
            {
                dt = selectliRecord(sblock, sline, stype, subtype, txtbarcode.Text);
            }
            else
            {
                dt = BindESDRecord(sblock, sline, stype, subtype, txtbarcode.Text);
            }
                    
            if (!this.txtstates.Text.Equals("ALL"))
            {
                DataRow[] arrDr = dt.Select("状态 <> '" + this.txtstates.Text + "'");
                foreach (DataRow dr in arrDr)
                    dt.Rows.Remove(dr);
            }
            this.databind.DataSource = dt;
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


        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "信息";
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

        private void btntoexcel_Click(object sender, EventArgs e)
        {

            DataTable dt = databind.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0)
                return;

            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridView);
            OpenFile(fileName);

        }
        private void btnsearchall_Click(object sender, EventArgs e)
        {
            databind.DataSource = null;
            DataTable dt = ic.AddNewTestESDList("所有查询", txtblock.Text == "" ? "" : txtblock.Text, txtline.Text == "" ? "" : txtline.Text, datetime.DateTime.ToString(),
                txttesttype.Text  == "" ? "" : txttesttype.Text, txttestsubitem.Text  == "" ? "" : txttestsubitem.Text, "", "", 0, "", "", txtbarcode.Text).Tables[0];
            if (dt.Rows.Count > 0)
            {
                gridView.Columns.Clear();
                databind.DataSource = null;
                databind.DataSource = dt;
            }
        }
        private void databind_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                //txtblock.Text = databind.CurrentRow.Cells["楼层"].Value.ToString();
                //txtline.SelectedValue = databind.CurrentRow.Cells["线别"].Value.ToString();
                //txttesttype.SelectedValue = databind.CurrentRow.Cells["测试类别"].Value.ToString();
                //txttestsubitem.SelectedValue = databind.CurrentRow.Cells["测试子项目"].Value.ToString();
                //txtLowervalue.Text = databind.CurrentRow.Cells["下限"].Value.ToString();
                //txtUppervalue.Text = databind.CurrentRow.Cells["上限"].Value.ToString();
            }
            catch
            {
            }
        }
        private void databind_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                //txtblock.Text = databind.CurrentRow.Cells["楼层"].Value.ToString();
                //txtline.SelectedValue = databind.CurrentRow.Cells["线别"].Value.ToString();
                //txttesttype.SelectedValue = databind.CurrentRow.Cells["测试类别"].Value.ToString();
                //txttestsubitem.SelectedValue = databind.CurrentRow.Cells["测试子项目"].Value.ToString();
                //txtLowervalue.Text = databind.CurrentRow.Cells["下限"].Value.ToString();
                //txtUppervalue.Text = databind.CurrentRow.Cells["上限"].Value.ToString();
            }
            catch
            {
            }
        }
        private void databind_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            //if (databind.Rows.Count <= 0)
            //    return;
            //try
            //{
            //    DataGridViewRow dgr = databind.Rows[e.RowIndex];
            //    if (dgr.Cells["状态"].Value.ToString() == "已过期")
            //    {
            //        dgr.DefaultCellStyle.BackColor = Color.Red;
            //    }
            //    else if (dgr.Cells["状态"].Value.ToString() == "快过期")
            //    {
            //        dgr.DefaultCellStyle.BackColor = Color.Yellow;

            //    }
            //    else if (dgr.Cells["测试结果"].Value.ToString() == "NG")
            //    {
            //        dgr.DefaultCellStyle.BackColor = Color.GreenYellow;

            //    }
            //}
            //catch
            //{
            //}
        }

        private void gridView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
        {
            if (gridView.RowCount < 1)
                return;
            try
            {       
                txtblock.Text = gridView.GetFocusedRowCellValue("楼层").ToString();
                txtline.Text = gridView.GetFocusedRowCellValue("线别").ToString();
                txttesttype.Text = gridView.GetFocusedRowCellValue("测试类别").ToString();
                txttestsubitem.Text = gridView.GetFocusedRowCellValue("测试子项目").ToString();
                txtLowervalue.Text = gridView.GetFocusedRowCellValue("下限").ToString();
                txtUppervalue.Text = gridView.GetFocusedRowCellValue("上限").ToString();
            }
            catch
            {
            }
        }

        private void gridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (gridView.RowCount < 1)
                return;
            try
            {
                txtblock.Text = gridView.GetFocusedRowCellValue("楼层").ToString();
                txtline.Text = gridView.GetFocusedRowCellValue("线别").ToString();
                txttesttype.Text = gridView.GetFocusedRowCellValue("测试类别").ToString();
                txttestsubitem.Text = gridView.GetFocusedRowCellValue("测试子项目").ToString();
                txtLowervalue.Text = gridView.GetFocusedRowCellValue("下限").ToString();
                txtUppervalue.Text = gridView.GetFocusedRowCellValue("上限").ToString();
            }
            catch
            {
            }

        }

        private void gridView_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {

            DataTable dt = databind.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            if (gridView.RowCount < 1)
                return;
            try
            {
                if (gridView.GetDataRow(e.RowHandle)["状态"].ToString() == "已过期")
                {
                    e.Appearance.BackColor = Color.Red;
                }
                else if (gridView.GetDataRow(e.RowHandle)["状态"].ToString() == "快过期")
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
                else if (gridView.GetDataRow(e.RowHandle)["测试结果"].ToString() == "NG")
                {
                    e.Appearance.BackColor = Color.GreenYellow;
                }
            }
            catch
            {
            }
        }

        private void cbOK_CheckedChanged(object sender, EventArgs e)
        {
            if (txttesttype.Text.Contains("离子") || txttesttype.Text.Contains("静电监控"))
            {
                if (cbOK.Checked == true)
                {
                    cbNG.Checked = false;
                }
                else
                {
                    cbNG.Checked = true;
                }
            }
        }

        private void cbNG_CheckedChanged(object sender, EventArgs e)
        {
            if (txttesttype.Text.Contains("离子") || txttesttype.Text.Contains("静电监控"))
            {
                if (cbNG.Checked == true)
                {
                    cbOK.Checked = false;
                }
                else
                {
                    cbOK.Checked = true;
                }
            }
        }
    }
}