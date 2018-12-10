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
using System.Data.OleDb;
using DX_QMS.Common;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using System.IO;
using System.Diagnostics;

namespace DX_QMS
{
    public partial class OQCTestProgSet : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public OQCTestProgSet()
        {
            InitializeComponent();
        }
        private void bind(string types, ComboBoxEdit com)
        {
            string ssql = "select Definevalue from OQC_TypeDefine where Definetype='"+types+" ' order by sort ";
            DataTable dt = DbAccess.SelectBySql(ssql).Tables[0];
            com.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
              com.Properties.Items.Add(row["Definevalue"]);
            }
        }
        private void OQCTestProgSet_Load(object sender, EventArgs e)
        {
            bind("客户", cboBcustomer);
            bind("测试项目",txttestitem);
            bind("检验工具",lUpcheckmethod);
            check_items.PageEnabled = false;
            rdoplan.SelectedIndex = 0;
            rdoMA.SelectedIndex = 3;
            rdoMI.SelectedIndex = 3;
            combIPC.Enabled = false;
            sBtnreset_Click(sender,e );
        }

        private void txtPN_Leave(object sender, EventArgs e)
        {
           
        }
        private void txtPN_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                gridControl.DataSource = null; 
                if (txtPN.Text == "") return;
                string sql = "select distinct item_num,item_desc from cux_wip_header_v v,CUX_WIP_OPERATIONS_V WO  where " +
                                      " WO.WIP_ENTITY_ID =v.WIP_ENTITY_ID and item_num ='" + txtPN.Text.Trim() + "'";  //and v.ORGANIZATION_CODE = '" + txtorg_id.SelectedValue.ToString() + "'";
                DataSet ds = DbAccess.SelectByOracle(sql);
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    txtPN.Text = ds.Tables[0].Rows[0]["item_num"].ToString();
                    string str = ds.Tables[0].Rows[0]["item_desc"].ToString();
                    lblinfo.Text = "";
                    System.Collections.ArrayList ltL = new System.Collections.ArrayList();
                    int index = 0;
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
                        txtmodel.Text = str.Substring(int.Parse(ltL[0].ToString()) + 1, int.Parse(ltR[0].ToString()) - int.Parse(ltL[0].ToString()) - 1);
                    }
                    catch
                    {
                    }

                }
                else
                {
                    lblinfo.Text = "提醒：" + txtPN.Text + "不存在";
                    lblinfo.ForeColor = Color.Red;
                    txtPN.Text = "";
                    return;
                }
            }
        }
        private void gridView_Click(object sender, EventArgs e)
        {

            if (bandedGridView.RowCount < 0)
            {
                MessageBox.Show("没有数据","提醒",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }
            DataTable de = gridControl.DataSource as DataTable;
            if ( de== null )
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } 
            cboBcustomer.Text = bandedGridView.GetFocusedRowCellValue("customer").ToString().Trim();
            txtPN.Text = bandedGridView.GetFocusedRowCellValue("PN").ToString().Trim();
            txtmodel.Text = bandedGridView.GetFocusedRowCellValue("model").ToString().Trim();
            /*
            string where = " where 1=1 ";
            if (!string.IsNullOrEmpty(cboBcustomer.Text.Trim()))
            {
                where += " and customer ='" + cboBcustomer.Text.Trim() + "' ";
            }
            if (!string.IsNullOrEmpty(txtPN.Text.Trim()))
            {
                where += " and PN ='" + txtPN.Text.Trim() + "' ";
            }
            string sql = @" select customer 客户,PN 编码,testitem 检查项目,checkmethod 检查方法,standardsequence 序号,teststandard 检验标准,checkMA MA主要缺陷,checkMI MI次要缺陷 from OQC_TestItem ";
            sql += where + " order by testitem,standardsequence ";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            gridList.DataSource = dt;  
            */

            /*
            txttestitem.Text = gridView.GetFocusedRowCellValue("检查项目").ToString();
            lUpcheckmethod.Text = gridView.GetFocusedRowCellValue("检查方法").ToString();
            txtteststandardsequence.Text = gridView.GetFocusedRowCellValue("序号").ToString();
            meEditteststandard.Text = gridView.GetFocusedRowCellValue("检验标准").ToString();
            string checkMA = gridView.GetFocusedRowCellValue("MA").ToString();

            if (checkMA == "是")
                raGroupMAMI.SelectedIndex = 0;
            string checkMI = gridView.GetFocusedRowCellValue("MI").ToString();
            if (checkMI == "是")
                raGroupMAMI.SelectedIndex = 1;

            rdoplan.SelectedIndex = getsampleplan(gridView.GetFocusedRowCellValue("抽样计划").ToString());
            rdoMA.SelectedIndex = getAQL(gridView.GetFocusedRowCellValue("MA主要缺陷").ToString());
            rdoMI.SelectedIndex = getAQL(gridView.GetFocusedRowCellValue("MI主要缺陷").ToString());
            string IPCA610F = gridView.GetFocusedRowCellValue("IPCA610F").ToString();
            if (IPCA610F == "")
                checkIPC.Enabled = false;
            else
            {
                checkIPC.Enabled = true;
                checkIPC.Checked = true;
                combIPC.Text = IPCA610F;
            }
            string customerstandard = gridView.GetFocusedRowCellValue("customerstandard").ToString();
            if (customerstandard == "")
                checkcustomer.Checked = false;
            else
                checkcustomer.Checked = true;
            string otherstandard = gridView.GetFocusedRowCellValue("otherstandard").ToString();
            if (otherstandard == "")
                checkother.Checked = false;
            else
                checkother.Checked = true;
            */




        }
        public string OQCTestProgSetAdd(string opertype, string customer, string PN, string model, string testitem, string checkmethod,int standardsequence,string teststandard,string checkMA,string checkMI, string sampleplan, string MA,double MAvalue,string MI, double MIvalue,string IPCA610F,string customerstandard,string otherstandard,string states)
        {
            SqlParameter[] para = new SqlParameter[20];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@customer", customer);
            para[2] = new SqlParameter("@PN", PN);
            para[3] = new SqlParameter("@model", model);
            para[4] = new SqlParameter("@testitem", testitem);
            para[5] = new SqlParameter("@checkmethod", checkmethod);
            para[6] = new SqlParameter("@standardsequence", standardsequence);
            para[7] = new SqlParameter("@teststandard", teststandard);
            para[8] = new SqlParameter("@checkMA", checkMA);
            para[9] = new SqlParameter("@checkMI", checkMI);
            para[10] = new SqlParameter("@sampleplan",sampleplan);
            para[11] = new SqlParameter("@MA", MA);
            para[12] = new SqlParameter("@MAvalue", MAvalue);
            para[13] = new SqlParameter("@MI", MI);
            para[14] = new SqlParameter("@MIvalue", MIvalue);
            para[15] = new SqlParameter("@IPCA610F",IPCA610F);
            para[16] = new SqlParameter("@customerstandard",customerstandard);
            para[17] = new SqlParameter("@otherstandard", otherstandard);
            para[18] = new SqlParameter("@states", states);
            para[19] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[19].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "OQC_TestProgSet", para);
            return para[19].Value.ToString();
        }

        public DataSet AddOQCTest_ProgSet(string opertype, string customer, string PN, string model, string testitem, string checkmethod, int standardsequence, string teststandard, string checkMA, string checkMI, string sampleplan, string MA, float MAvalue, string MI, float MIvalue, string IPCA610F, string customerstandard, string otherstandard, string states)
        {
            SqlParameter[] para = new SqlParameter[20];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@customer", customer);
            para[2] = new SqlParameter("@PN", PN);
            para[3] = new SqlParameter("@model", model);
            para[4] = new SqlParameter("@testitem", testitem);
            para[5] = new SqlParameter("@checkmethod", checkmethod);
            para[6] = new SqlParameter("@standardsequence", standardsequence);
            para[7] = new SqlParameter("@teststandard", teststandard);
            para[8] = new SqlParameter("@checkMA", checkMA);
            para[9] = new SqlParameter("@checkMI", checkMI);
            para[10] = new SqlParameter("@sampleplan", sampleplan);
            para[11] = new SqlParameter("@MA", MA);
            para[12] = new SqlParameter("@MAvalue", MAvalue);
            para[13] = new SqlParameter("@MI", MI);
            para[14] = new SqlParameter("@MIvalue", MIvalue);
            para[15] = new SqlParameter("@IPCA610F", IPCA610F);
            para[16] = new SqlParameter("@customerstandard", customerstandard);
            para[17] = new SqlParameter("@otherstandard", otherstandard);
            para[18] = new SqlParameter("@states", states);
            para[19] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[19].Direction = ParameterDirection.Output;
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "OQC_TestProgSet", para);
        }

        public string OQCTestItemAdd(string opertype,string customer,string PN, string testitem, string checkmethod, int standardsequence, string teststandard ,string checkMA ,string checkMI )
        {
            SqlParameter[] para = new SqlParameter[10];  
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@customer", customer);
            para[2] = new SqlParameter("@PN", PN);
            para[3] = new SqlParameter("@testitem", testitem);
            para[4] = new SqlParameter("@checkmethod", checkmethod);
            para[5] = new SqlParameter("@standardsequence", standardsequence);
            para[6] = new SqlParameter("@teststandard", teststandard);
            para[7] = new SqlParameter("@checkMA", checkMA);
            para[8] = new SqlParameter("@checkMI", checkMI);
            para[9] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[9].Direction = ParameterDirection.Output;
            DbAccess.ExecuteNonQuery(CommandType.StoredProcedure, "OQC_InsertTestItem", para);
            return para[9].Value.ToString();
        }

        public DataSet OQCTestItemReasch(string opertype, string customer, string PN, string testitem, string checkmethod, int standardsequence, string teststandard, string checkMA, string checkMI)
        {
            SqlParameter[] para = new SqlParameter[10];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@customer", customer);
            para[2] = new SqlParameter("@PN", PN);
            para[3] = new SqlParameter("@testitem", testitem);
            para[4] = new SqlParameter("@checkmethod", checkmethod);
            para[5] = new SqlParameter("@standardsequence", standardsequence);
            para[6] = new SqlParameter("@teststandard", teststandard);
            para[7] = new SqlParameter("@checkMA", checkMA);
            para[8] = new SqlParameter("@checkMI", checkMI);
            para[9] = new SqlParameter("@msg", SqlDbType.VarChar, 50);
            para[9].Direction = ParameterDirection.Output;
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "OQC_InsertTestItem", para);
        }

        string Sampleplan(int i)
        {   
            string sample = "";
            switch (i)
               {
                    case 0: sample = "C=0";
                        break;
                    case 1: sample = "ISO2859-1";
                        break;
                    case 2: sample = "全检";
                        break;
                    default:sample = "";
                        break;
                }
            return sample;
        }

        int getsampleplan(string sampleplan)
        {
            int type;
            if (sampleplan == "C=0")
                type = 0;
            else if (sampleplan == "ISO2859-1")
                type = 1;
            else 
                type = 2;
            return type;
        }
        string AQL(int i)
        {
            string MA = "";
            switch (i)
            {
                case 0:
                    MA = "AQL=0.65";
                    break;
                case 1:
                    MA = "AQL=0.40";
                    break;
                case 2:
                    MA = "AQL=0.01";
                    break;
                default:
                    MA = "其他";
                    break;
            }
            return MA;
        }

        int getAQL(string AQLtype)
        {
            int type;
            if (AQLtype == "AQL=0.65")
                type = 0;
            else if (AQLtype == "AQL=0.40")
                type = 1;
            else if (AQLtype == "AQL=0.01")
                type = 2;
            else
                type = 3;
            return type;
        }
        double AQLvalue (int i)
        {
            double value = 0;
            switch (i)
            {
                case 0:
                    value = 0.65;
                    break;
                case 1:
                    value = 0.40;
                    break;
                case 2:
                    value = 0.01;
                    break;
                default:
                    value = 0;
                    break;
            }
            return value;
        }

        void AddRow(DevExpress.XtraGrid.Views.Grid.GridView view)
        {
            int prevDataRowIndex = view.GetFocusedDataSourceRowIndex();
            view.AddNewRow();
            if (view.GroupCount >= 0 && prevDataRowIndex >= 0)
            {
                foreach (DevExpress.XtraGrid.Columns.GridColumn groupColumn in view.GroupedColumns)
                {
                    object val = view.GetRowCellValue(prevDataRowIndex, groupColumn);
                    view.SetFocusedRowCellValue(groupColumn, val);
                }
                view.UpdateCurrentRow();
            }
            view.ShowEditor();
        }


         public static DataTable CreateTable()
        {
            DataTable namesTable = new DataTable("information");
            DataColumn customerColumn = new DataColumn();
            customerColumn.DataType = System.Type.GetType("System.String");
            customerColumn.ColumnName = "customer";
            namesTable.Columns.Add(customerColumn);

            DataColumn PNColumn = new DataColumn();
            PNColumn.DataType = System.Type.GetType("System.String");
            PNColumn.ColumnName = "PN";
            namesTable.Columns.Add(PNColumn);

            DataColumn modelColumn = new DataColumn();
            modelColumn.DataType = System.Type.GetType("System.String");
            modelColumn.ColumnName = "model";
            namesTable.Columns.Add(modelColumn);
            return namesTable;
        }

          DataTable table = CreateTable();

        private void sBtnadd_Click(object sender, EventArgs e)
        {
            /*
            if (cboBcustomer.Text == "" ||(cboBcustomer.Text == "" && txtPN.Text !=""))
            {
                MessageBox.Show("请输入客户或者P/N号","停止",MessageBoxButtons.OK,MessageBoxIcon.Stop);
                return;
            }
            DataRow row;
            row = table.NewRow();
            row["customer"] = cboBcustomer.Text;
            row["PN"] = txtPN.Text ;
            row["model"] = txtmodel.Text ;
            table.Rows.Add(row);
            gridControl.DataSource = table;
            */
            string customer = "", PN = "";
            string sql = "";
            customer = cboBcustomer.Text.Trim();
            PN = txtPN.Text.Trim();
            if (customer != "" && PN !="")
            {
                sql = @"select customer 客户,PN 编码,model 机型,testitem 检验项目,checkmethod 检查方法,standardsequence 检验标准序列,teststandard 检验标准,checkMA 主要缺陷,checkMI 次要缺陷,sampleplan 抽样计划,MA MA主要缺陷,MAvalue MA值,MI MI次要缺陷,MIvalue MI值,IPCA610F IPCA610F,customerstandard 客户检验标准规范,otherstandard 其他,SNstates 是否扫描SN号 from OQCTestProgSet ";
                sql = sql + "where customer= '"+ customer + "' and PN='"+ PN + "'";
            }
            if (customer != ""&& PN == "")
            {
                sql = @"select customer 客户,PN 编码,model 机型,testitem 检验项目,checkmethod 检查方法,standardsequence 检验标准序列,teststandard 检验标准,checkMA 主要缺陷,checkMI 次要缺陷,sampleplan 抽样计划,MA MA主要缺陷,MAvalue MA值,MI MI次要缺陷,MIvalue MI值,IPCA610F IPCA610F,customerstandard 客户检验标准规范,otherstandard 其他,SNstates 是否扫描SN号 from OQCTestProgSet ";
                sql = sql + "where customer= '" + customer + "' and PN =''";
            }
            if (customer == "" && PN == "")
            {
                sql = @"select customer 客户,PN 编码,model 机型,testitem 检验项目,checkmethod 检查方法,standardsequence 检验标准序列,teststandard 检验标准,checkMA 主要缺陷,checkMI 次要缺陷,sampleplan 抽样计划,MA MA主要缺陷,MAvalue MA值,MI MI次要缺陷,MIvalue MI值,IPCA610F IPCA610F,customerstandard 客户检验标准规范,otherstandard 其他,SNstates 是否扫描SN号 from OQCTestProgSet ";
                sql = sql + "where customer='' and PN=''";
            }
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt != null && dt.Rows.Count > 0)
                gridControl.DataSource = dt;
            else
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                gridControl.DataSource = null;
            }
        }
        private void sBtnpreview_Click(object sender, EventArgs e)
        { 

            string customer = "", PN = "", model ="",testitem = "", checkmethod = "", checkMA = "否", checkMI = "否";
            string sampleplan = "", MA = "", MI = "", IPCA610F = "", customerstandard = "", otherstandard = "", states= "";
            double MAvalue, MIvalue;
            int standardsequence = 0;
            if (!int.TryParse(txtteststandardsequence.Text, out standardsequence))
            {
                standardsequence = 0;
            }
            if (standardsequence <= 0)
            {
                standardsequence = 0;
            }
            if (raGroupMAMI.SelectedIndex == 0)
            {
                checkMA = "是";
            }
            if (raGroupMAMI.SelectedIndex == 1)
            {
                checkMI = "是";
            }
            customer = cboBcustomer.Text.Trim();
            PN = txtPN.Text.Trim();
            model = txtmodel.Text;
            testitem = txttestitem.Text.Trim();
            sampleplan = Sampleplan(rdoplan.SelectedIndex);
            MA = AQL(rdoMA.SelectedIndex);
            MI = AQL(rdoMI.SelectedIndex);
            MAvalue = AQLvalue(rdoMA.SelectedIndex);
            MIvalue = AQLvalue(rdoMI.SelectedIndex);  
            if (checkIPC.Checked)
                IPCA610F = combIPC.Text;
            if (checkcustomer.Checked)
                customerstandard = "客户检验标准规范";
            if (checkother.Checked)
                otherstandard = "其他";

            try
            {
                int sheetCount = 1;
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = true;
                object missing = System.Reflection.Missing.Value;
                string templetFile = Environment.CurrentDirectory + @"\ReportFolder\OQCreport.xlsx";
                Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(templetFile, missing, true, missing, missing, missing,
                                                              missing, missing, missing, missing, missing, missing, missing, missing, missing);
                Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets.get_Item(1);

                for (int i = 1; i < sheetCount; i++)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(i)).Copy(missing, workBook.Worksheets[i]);
                }
                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);
                if (sheet == null)
                    return;

                sheet.Cells.get_Range("C4").Value = customer;
                sheet.Cells.get_Range("C5").Value = model;
                sheet.Cells.get_Range("C7").Value = PN;
                if (sampleplan == "C=0")
                    sheet.Cells.get_Range("N3").Value = "√C=0";
                else if (sampleplan == "ISO2859-1")
                    sheet.Cells.get_Range("O3").Value = "√ISO2859-1一般II级";
                else
                    sheet.Cells.get_Range("Q3").Value = "√全检";

                if (MA == "AQL=0.65")
                    sheet.Cells.get_Range("N4").Value = "√0.65";
                else if (MA == "AQL=0.4")
                    sheet.Cells.get_Range("O4").Value = "√0.4";
                else if (MA == "AQL=0.01")
                    sheet.Cells.get_Range("P4").Value = "√0.01";
                else
                    sheet.Cells.get_Range("Q4").Value = "√其他";

                if (MI == "AQL=0.65")
                    sheet.Cells.get_Range("N5").Value = "√0.65";
                else if (MI == "AQL=0.4")
                    sheet.Cells.get_Range("O5").Value = "√0.4";
                else if (MI == "AQL=0.01")
                    sheet.Cells.get_Range("P5").Value = "√0.01";
                else
                    sheet.Cells.get_Range("Q5").Value = "√其他";

                if (IPCA610F == "I级")
                    sheet.Cells.get_Range("O6").Value = "√I级";
                if (IPCA610F == "II级")
                    sheet.Cells.get_Range("P6").Value = "√II级";
                if (IPCA610F == "III级")
                    sheet.Cells.get_Range("Q6").Value = "√III级";
                if (customerstandard == "客户检验标准规范")
                    sheet.Cells.get_Range("N7").Value = "√客户检验标准规范";
                if (otherstandard == "其他")
                    sheet.Cells.get_Range("N8").Value = "√其他";

                DataTable copy = gridControltosql.DataSource as DataTable;
                if (copy == null)
                {
                    MessageBox.Show("没有检验项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                int count = copy.Rows.Count;
                if (count < 0)
                {
                    MessageBox.Show("没有检查项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                for (int n = 0, m = 11; n < count; n++, m++)
                {
                    sheet.Cells[m, 1] = copy.Rows[n]["testitem"].ToString();
                    sheet.Cells[m, 3] = copy.Rows[n]["standardsequence"].ToString();
                    sheet.Cells[m, 4] = copy.Rows[n]["teststandard"].ToString();
                    sheet.Cells[m, 13] = copy.Rows[n]["checkmethod"].ToString();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("预览失败！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void sBtndelete_Click(object sender, EventArgs e)
        {
            
            int i = bandedGridView.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选中要删除的维护项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            if (bandedGridView.GetSelectedRows().Length < 0)
                return;
            /*
            string customer = cboBcustomer.Text;
            string PN = txtPN.Text;
            string model = txtmodel.Text;
            string testitem = txttestitem.Text;
            if (testitem == "")
            {
                MessageBox.Show("请输入检查项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string checkmethod = lUpcheckmethod.Text;
            int standardsequence = 0;
            if (!int.TryParse(txtteststandardsequence.Text, out standardsequence))
            {
                standardsequence = 0;
                lblsequence.Text = "请输入正确的数字";
                lblsequence.ForeColor = Color.Red;
                return;
            }
            if (standardsequence <= 0)
            {
                standardsequence = 0;
                lblsequence.Text = "请输入正确的数字";
                lblsequence.ForeColor = Color.Red;
                return;
            }
            string teststandard = meEditteststandard.Text;
            string checkMA = "否", checkMI = "否";
            string sampleplan = "", MA = "", MI = "";
            double MAvalue, MIvalue;
            string IPCA610F = "", customerstandard = "", otherstandard = "", states = "";
            if (raGroupMAMI.SelectedIndex == 0)
            {
                checkMA = "是";
            }
            if (raGroupMAMI.SelectedIndex == 1)
            {
                checkMI = "是";
            }
            if (checkIPC.Checked)
                IPCA610F = combIPC.Text;
            if (checkcustomer.Checked)
                customerstandard = "客户检验标准规范";
            if (checkother.Checked)
                otherstandard = "其他";
            sampleplan = Sampleplan(rdoplan.SelectedIndex);
            MA = AQL(rdoMA.SelectedIndex);
            MI = AQL(rdoMI.SelectedIndex);
            MAvalue = AQLvalue(rdoMA.SelectedIndex);
            MIvalue = AQLvalue(rdoMI.SelectedIndex);
            */


            //string customer = bandedGridView.GetFocusedRowCellValue("客户").ToString();
            //string PN = bandedGridView.GetFocusedRowCellValue("编码").ToString();
            //string model = bandedGridView.GetFocusedRowCellValue("机型").ToString();
            //string testitem = bandedGridView.GetFocusedRowCellValue("检验项目").ToString();
            //string checkmethod = bandedGridView.GetFocusedRowCellValue("检查方法").ToString();
            //string standardsequence = bandedGridView.GetFocusedRowCellValue("检验标准序列").ToString();
            //string teststandard = bandedGridView.GetFocusedRowCellValue("检验标准").ToString();
            //string checkMA = bandedGridView.GetFocusedRowCellValue("主要缺陷").ToString();
            //string checkMI = bandedGridView.GetFocusedRowCellValue("次要缺陷").ToString();
            //string sampleplan = bandedGridView.GetFocusedRowCellValue("抽样计划").ToString();
            //string MA = bandedGridView.GetFocusedRowCellValue("MA主要缺陷").ToString();
            //string MI = bandedGridView.GetFocusedRowCellValue("MI次要缺陷").ToString();
            //string IPCA610F = bandedGridView.GetFocusedRowCellValue("IPCA610F").ToString();
            //string customerstandard = bandedGridView.GetFocusedRowCellValue("客户检验标准规范").ToString();
            //string otherstandard = bandedGridView.GetFocusedRowCellValue("其他").ToString();
            //string SNstates = bandedGridView.GetFocusedRowCellValue("是否扫描SN号").ToString();
            //string MAvalue = bandedGridView.GetFocusedRowCellValue("MA值").ToString();
            //string MIvalue = bandedGridView.GetFocusedRowCellValue("MI值").ToString();  dr["客户"].ToString()
            int count = 0;
            for (int k = bandedGridView.GetSelectedRows().Length; k > 0; k--)
            {
                DataRow dr = bandedGridView.GetDataRow(bandedGridView.GetSelectedRows()[k - 1]);
                string customer = dr["客户"].ToString();
                string PN = dr["编码"].ToString();
                string model = dr["机型"].ToString();
                string testitem = dr["检验项目"].ToString();
                string checkmethod = dr["检查方法"].ToString();
                string standardsequence = dr["检验标准序列"].ToString();
                string teststandard = dr["检验标准"].ToString();
                string checkMA = dr["主要缺陷"].ToString();
                string checkMI = dr["次要缺陷"].ToString();
                string sampleplan = dr["抽样计划"].ToString();
                string MA = dr["MA主要缺陷"].ToString();
                string MI = dr["MI次要缺陷"].ToString();
                string IPCA610F = dr["IPCA610F"].ToString();
                string customerstandard = dr["客户检验标准规范"].ToString();
                string otherstandard = dr["其他"].ToString();
                string SNstates = dr["是否扫描SN号"].ToString();
                string MAvalue = dr["MA值"].ToString();
                string MIvalue = dr["MI值"].ToString();

                string flag = OQCTestProgSetAdd("删除", customer, PN, model, testitem, checkmethod, int.Parse(standardsequence), teststandard, checkMA, checkMI, sampleplan, MA, double.Parse(MAvalue), MI, double.Parse(MIvalue), IPCA610F, customerstandard, otherstandard, SNstates);

                if (flag.Contains("删除成功"))
                {
                    count = count + 1;
                }

            }

            MessageBox.Show("成功删除"+ count + "条", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            sBtnadd_Click(sender, e);

            //if (flag == "删除成功")
            //{
            //    MessageBox.Show(flag, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    sBtnadd_Click(sender, e);
            //}
            //else
            //    MessageBox.Show(flag, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);


            /*
            int i = gridView.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选中需要删除的信息", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            cboBcustomer.Text = gridView.GetFocusedRowCellValue("customer").ToString();
            txtPN.Text = gridView.GetFocusedRowCellValue("PN").ToString();
            txtmodel.Text = gridView.GetFocusedRowCellValue("model").ToString();
            int n;
            for (n = table.Rows.Count - 1; n >= 0; n--)
            {
                if ( (table.Rows[n]["customer"].ToString () == cboBcustomer.Text) && (table.Rows[n]["PN"].ToString() == txtPN.Text) && (table.Rows[n]["model"].ToString() == txtmodel.Text))
                {
                    table.Rows[n].Delete();
                }          
            }
            gridControl.DataSource = table;
            */
        }

        private void sBtnsubmit_Click(object sender, EventArgs e)   
        {

            if (DialogResult.Cancel  == MessageBox.Show("是否提交检验维护信息", "询问", MessageBoxButtons.OKCancel, MessageBoxIcon.Question))
            {
                return;
            }
            string customer = cboBcustomer.Text;
            string PN = txtPN.Text;
            string model = txtmodel.Text;
            string sampleplan = "", MA = "", MI = "";
            double MAvalue, MIvalue;
            string IPCA610F = "", customerstandard = "", otherstandard = "", SNstates = "否";
            if (checkIPC.Checked)
                IPCA610F = combIPC.Text;
            if (checkcustomer.Checked)
                customerstandard = "客户检验标准规范";
            if (checkother.Checked)
                otherstandard = "其他";
            if (checkSN.Checked)
                SNstates = "是";
            sampleplan = Sampleplan(rdoplan.SelectedIndex);
            MA = AQL(rdoMA.SelectedIndex);
            MI = AQL(rdoMI.SelectedIndex);
            MAvalue = AQLvalue(rdoMA.SelectedIndex);
            MIvalue = AQLvalue(rdoMI.SelectedIndex);
            DataTable copy = gridControltosql.DataSource as DataTable;
            if (copy == null)
            {
                MessageBox.Show("没有检验项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            int count = copy.Rows.Count;
            int countsuccess = 0;
            if (count <0)
            {
                MessageBox.Show("没有检验项目","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;

            }
            for (int n = 0; n < count; n++)
            {
                string testitem = copy.Rows[n]["testitem"].ToString();
                string standardsequence = copy.Rows[n]["standardsequence"].ToString();
                string teststandard = copy.Rows[n]["teststandard"].ToString();
                string checkmethod = copy.Rows[n]["checkmethod"].ToString();  
                string checkMA= copy.Rows[n]["checkMA"].ToString();
                string checkMI= copy.Rows[n]["checkMI"].ToString();
                string flag = OQCTestProgSetAdd("添加", customer, PN, model, testitem, checkmethod, int.Parse(standardsequence), teststandard, checkMA, checkMI, sampleplan, MA, MAvalue, MI, MIvalue, IPCA610F, customerstandard, otherstandard,SNstates);
                if (flag == "添加成功")
                {
                   countsuccess++; 
                }
            }
            MessageBox.Show("提交成功"+countsuccess+"条,失败"+(count-countsuccess)+"条" ,"提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);

        }

        private void sBtnrearsh_Click(object sender, EventArgs e)
        {
            /*
            string customer = cboBcustomer.Text;
            string PN = txtPN.Text;
            string testitem = txttestitem.Text;
            if (testitem == "")
            {
                MessageBox.Show("请输入检验项目", "警告", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string checkmethod = lUpcheckmethod.Text;
            int standardsequence = 0;
            if (!int.TryParse(txtteststandardsequence.Text, out standardsequence))
            {
                standardsequence = 0;
            }
            if (standardsequence <= 0)
            {
                standardsequence = 0;
            }
            string checkMA = "否", checkMI = "否";
            if (checkEditMA.Checked)
                checkMA = "是";  
            if (checkEditMI.Checked)
                checkMI = "是";
            DataSet det = OQCTestItemReasch("查询",customer, PN, testitem,checkmethod,standardsequence,"", checkMA, checkMI);
            DataTable dt = det.Tables[0];
            if (dt != null && dt.Rows.Count > 0)
                gridList.DataSource = dt;
            else
                MessageBox.Show("没有相关的记录","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
             */



            string where = " where 1=1 ";
            string customer="", PN="", testitem = "", checkmethod = "", checkMA = "否", checkMI = "否";
            int standardsequence = 0;
            if (!int.TryParse(txtteststandardsequence.Text, out standardsequence))
            {
                standardsequence = 0;
            }
            if (standardsequence <= 0)
            {
                standardsequence = 0;
            }
            if (raGroupMAMI.SelectedIndex == 0)
                checkMA = "是";
            if (raGroupMAMI.SelectedIndex == 1)
                checkMI = "是";

            customer = cboBcustomer.Text.Trim();
            PN = txtPN.Text.Trim();
            testitem = txttestitem.Text.Trim();
            checkmethod = lUpcheckmethod.Text.Trim();

            if (!string.IsNullOrEmpty(customer))
            {
                where += " and customer ='" + customer + "' ";
            }
            if (!string.IsNullOrEmpty(PN))
            {
                where += " and PN ='" + PN + "' ";
            }
            if (!string.IsNullOrEmpty(testitem))
            {
                where += " and testitem ='" + testitem + "' ";
            }
            if (!string.IsNullOrEmpty(checkmethod))
            {
                where += " and checkmethod like '%" + checkmethod + "%' ";  
            }
            if (!string.IsNullOrEmpty(checkMA))
            {
                where += " and checkMA ='" + checkMA + "' ";
            }
            if (!string.IsNullOrEmpty(checkMI))
            {
                where += " and checkMI ='" + checkMI + "' ";
            }
            if (standardsequence > 0)
            {
                where += " and standardsequence =" + standardsequence;   
            }
            string sql = @" select customer 客户,PN 编码,testitem 检查项目,checkmethod 检查方法,standardsequence 序号,teststandard 检验标准,checkMA MA主要缺陷,checkMI MI次要缺陷 from OQC_TestItem ";
            sql += where + " order by testitem,standardsequence ";  
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0]; 
            if (dt != null && dt.Rows.Count > 0)
                gridList.DataSource = dt;
            else
                MessageBox.Show("没有相关的记录", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void sBtnaddcheck_Click(object sender, EventArgs e)
        {
            lblsequence.Text = "";

            string customer, PN;
            customer= cboBcustomer.Text;
            PN = txtPN.Text;
            string testitem = txttestitem.Text;
            if (testitem == "")
            {
                MessageBox.Show("请输入检验项目","警告",MessageBoxButtons.OK ,MessageBoxIcon.Stop);
                return;
            }
            string checkmethod = lUpcheckmethod.Text;
            if (checkmethod == "")
            {
                MessageBox.Show("请输入检验方法", "警告", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            int standardsequence = 0;
            if (!int.TryParse(txtteststandardsequence.Text, out standardsequence))
            {
                txtteststandardsequence.Text = "";
                lblsequence.Text = "请输入正确的数字";
                lblsequence.ForeColor = Color.Red;
                return;
            }
            if (standardsequence <= 0)
            {
                txtteststandardsequence.Text = "";
                lblsequence.Text = "请输入正确的数字";
                lblsequence.ForeColor = Color.Red;
                return;
            }
            string teststandard = meEditteststandard.Text;

            string checkMA = "否", checkMI = "否";
            if (raGroupMAMI.SelectedIndex == 0)
            {
                checkMA = "是";
            }
            if (raGroupMAMI.SelectedIndex == 1)
            {
                checkMI = "是";
            }
            string flag = OQCTestItemAdd("新增",customer,PN,testitem,checkmethod,standardsequence, teststandard, checkMA, checkMI);
            if (flag == "添加成功")
                MessageBox.Show(flag, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(flag, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void sBtnsendaudit_Click(object sender, EventArgs e)
        {
            // DataRow dr = this.gridListView.GetDataRow(gridListView.FocusedRowHandle);  //dr.IsNull(3)
            int i = gridListView.FocusedRowHandle;
            if ( i < 0 )
            {
                MessageBox.Show("请选中需要审核的检验项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            int standardsequence = 0;
            if (!int.TryParse(txtteststandardsequence.Text, out standardsequence))
            {
                standardsequence = 0;
            }
            if (standardsequence <= 0)
            {
                standardsequence = 0;
            }
            string checkMA = gridListView.GetFocusedRowCellValue("MA").ToString();
            string checkMI = gridListView.GetFocusedRowCellValue("MI").ToString();
            //string flag = OQCTestItemAdd("审核", cboBcustomer.Text, txtPN.Text, txttestitem.Text, lUpcheckmethod.Text, standardsequence, meEditteststandard.Text, checkMA, checkMI);
            //if (flag == "审核失败")
            //    MessageBox.Show(flag, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //else
            //    MessageBox.Show("审核成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void sBtnremove_Click(object sender, EventArgs e)
        {
            int i = gridListView.FocusedRowHandle;
            if (i < 0)
             { 
               MessageBox.Show("请选中需要移除的检验项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
             }  
            string customer = cboBcustomer.Text.Trim();
            string PN = txtPN.Text.Trim();
            string testitem = txttestitem.Text.Trim();
            if (testitem == "")
            {
                MessageBox.Show("检验项目不能为空","警告",MessageBoxButtons.OK ,MessageBoxIcon.Warning);
                return;
            }
            string checkmethod = lUpcheckmethod.Text.Trim();
            if (checkmethod == "")
            {
                MessageBox.Show("检验方法不能为空", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            int standardsequence = 0;
            if (!int.TryParse(txtteststandardsequence.Text, out standardsequence))
            {
                standardsequence = 0;
            }
            if (standardsequence <= 0)
            {
                standardsequence = 0;
            }
            string checkMA = "否", checkMI = "否";
            if (raGroupMAMI.SelectedIndex == 0)
            {
                checkMA = "是";
            }
            if (raGroupMAMI.SelectedIndex == 1)
            {
                checkMI = "是";
            }
            string flag = OQCTestItemAdd("移除",customer,PN,testitem,checkmethod, standardsequence, meEditteststandard.Text,checkMA,checkMI);
            if (flag == "不存在这条记录")
                MessageBox.Show(flag,"警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
                MessageBox.Show("删除成功", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            gridList.DataSource = null;
            sBtnreset_Click(sender, e);
            sBtnrearsh_Click(sender,e);
        }

        private void gridListView_Click(object sender, EventArgs e)
        {            
            cboBcustomer.Text = gridListView.GetFocusedRowCellValue("客户").ToString();
            txtPN.Text = gridListView.GetFocusedRowCellValue("编码").ToString();
            txttestitem.Text = gridListView.GetFocusedRowCellValue("检查项目").ToString();
            lUpcheckmethod.Text = gridListView.GetFocusedRowCellValue("检查方法").ToString();
            txtteststandardsequence.Text = gridListView.GetFocusedRowCellValue("序号").ToString();
            meEditteststandard.Text = gridListView.GetFocusedRowCellValue("检验标准").ToString();
            string checkMA = gridListView.GetFocusedRowCellValue("MA主要缺陷").ToString();
            if (checkMA == "是")
                raGroupMAMI.SelectedIndex = 0;
            string checkMI = gridListView.GetFocusedRowCellValue("MI次要缺陷").ToString();
            if (checkMI == "是")
                raGroupMAMI.SelectedIndex = 1;
        }

        private void sBtnreset_Click(object sender, EventArgs e)
        {
            cboBcustomer.Text = "";
            txtPN.Text = "";
            txttestitem.Text = "";
            lUpcheckmethod.Text = "";
            txtteststandardsequence.Text = "";
            meEditteststandard.Text = "";
        }

        private void checkIPC_CheckedChanged(object sender, EventArgs e)
        {
            if (checkIPC.Checked)
            {
                combIPC.Enabled = true;
                combIPC.SelectedIndex = 0;
            }
            else
                combIPC.Enabled = false;
        }

        private void btnaudittest_Click(object sender, EventArgs e)
        {
            int i = bandedGridView.FocusedRowHandle;
            if (i < 0)
            {
                MessageBox.Show("请选中需要修改的项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            /*
            string customer = cboBcustomer.Text;
            string PN = txtPN.Text;
            string model = txtmodel.Text;
            string testitem = txttestitem.Text;
            if (testitem == "")
            {
                MessageBox.Show("请输入检查项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string checkmethod = lUpcheckmethod.Text;
            int standardsequence = 0;
            if (!int.TryParse(txtteststandardsequence.Text, out standardsequence))
            {
                standardsequence = 0;
                lblsequence.Text = "请输入正确的数字";
                lblsequence.ForeColor = Color.Red;
                return;
            }
            if (standardsequence <= 0)
            {
                standardsequence = 0;
                lblsequence.Text = "请输入正确的数字";
                lblsequence.ForeColor = Color.Red;
                return;
            }
            string teststandard = meEditteststandard.Text;
            string checkMA = "否", checkMI = "否";
            string sampleplan = "", MA = "", MI = "";
            double MAvalue, MIvalue;
            string IPCA610F = "", customerstandard = "", otherstandard = "", states = "";
            if (raGroupMAMI.SelectedIndex == 0)
            {
                checkMA = "是";
            }
            if (raGroupMAMI.SelectedIndex == 1)
            {
                checkMI = "是";
            }
            if (checkIPC.Checked)
                IPCA610F = combIPC.Text;
            if (checkcustomer.Checked)
                customerstandard = "客户检验标准规范";
            if (checkother.Checked)
                otherstandard = "其他";
            sampleplan = Sampleplan(rdoplan.SelectedIndex);
            MA = AQL(rdoMA.SelectedIndex);
            MI = AQL(rdoMI.SelectedIndex);
            MAvalue = AQLvalue(rdoMA.SelectedIndex);
            MIvalue = AQLvalue(rdoMI.SelectedIndex);         
            string flag = OQCTestProgSetAdd("审核", customer, PN, model, testitem, checkmethod, standardsequence, teststandard, checkMA, checkMI, sampleplan, MA, MAvalue, MI, MIvalue, IPCA610F, customerstandard, otherstandard, "");
            if (flag == "审核成功")
                MessageBox.Show(flag, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(flag, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                */
            string customer = bandedGridView.GetFocusedRowCellValue("客户").ToString();
            string PN = bandedGridView.GetFocusedRowCellValue("编码").ToString();
            string model = bandedGridView.GetFocusedRowCellValue("机型").ToString();
            string testitem = bandedGridView.GetFocusedRowCellValue("检验项目").ToString();
            string checkmethod = bandedGridView.GetFocusedRowCellValue("检查方法").ToString();
            string standardsequence = bandedGridView.GetFocusedRowCellValue("检验标准序列").ToString();
            string teststandard = bandedGridView.GetFocusedRowCellValue("检验标准").ToString();
            string checkMA = bandedGridView.GetFocusedRowCellValue("主要缺陷").ToString();
            string checkMI = bandedGridView.GetFocusedRowCellValue("次要缺陷").ToString();
            //string sampleplan = bandedGridView.GetFocusedRowCellValue("抽样计划").ToString();
            //string MA = bandedGridView.GetFocusedRowCellValue("MA主要缺陷").ToString();
            //string MI = bandedGridView.GetFocusedRowCellValue("MI次要缺陷").ToString();
            //string IPCA610F = bandedGridView.GetFocusedRowCellValue("IPCA610F").ToString();
            //string customerstandard = bandedGridView.GetFocusedRowCellValue("客户检验标准规范").ToString();
            //string otherstandard = bandedGridView.GetFocusedRowCellValue("其他").ToString();
            //string MAvalue = bandedGridView.GetFocusedRowCellValue("MA值").ToString();
            //string MIvalue = bandedGridView.GetFocusedRowCellValue("MI值").ToString();
            string sampleplan = Sampleplan(rdoplan.SelectedIndex);
            string MA = AQL(rdoMA.SelectedIndex);
            string MI = AQL(rdoMI.SelectedIndex);
            string IPCA610F = "";
            string customerstandard = "";
            string otherstandard = "";
            string SNstates = "否";
            if (checkIPC.Checked)
                IPCA610F = combIPC.Text;
            if (checkcustomer.Checked)
                customerstandard = "客户检验标准规范";
            if (checkother.Checked)
                otherstandard = "其他";
            if (checkSN.Checked)
                SNstates = "是";

            double  MAvalue = AQLvalue(rdoMA.SelectedIndex);            
            double  MIvalue = AQLvalue(rdoMI.SelectedIndex);

            if (customer != "" && PN != "")
            {
                string sql = @" update OQCTestProgSet set sampleplan ='" + sampleplan + "', MA ='" + MA + "', MAvalue=" + MAvalue + ", MI ='" + MI + "', MIvalue =" + MIvalue + ",IPCA610F = '" + IPCA610F + "',customerstandard ='" + customerstandard + "', otherstandard = '" + otherstandard + "',SNstates = '"+ SNstates+ "' where customer ='" + customer + "' and PN ='" + PN + "'";
                bool flag = DbAccess.ExecuteSql(sql);
                if (flag)
                {
                    MessageBox.Show("修改成功！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    sBtnadd_Click(sender, e);
                }
                else
                {
                    MessageBox.Show("修改失败！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (customer != "" && PN == "")
            {
                string sql = @" update OQCTestProgSet set sampleplan ='" + sampleplan + "', MA ='" + MA + "', MAvalue=" + MAvalue + ", MI ='" + MI + "', MIvalue =" + MIvalue + ",IPCA610F = '" + IPCA610F + "',customerstandard ='" + customerstandard + "', otherstandard = '" + otherstandard + "',SNstates = '"+ SNstates + "' where customer ='" + customer + "' and PN =''";
                bool flag = DbAccess.ExecuteSql(sql);
                if (flag)
                {
                    MessageBox.Show("修改成功！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    sBtnadd_Click(sender, e);
                }
                else
                {
                    MessageBox.Show("修改失败！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (customer == "" && PN == "")
            {
                string sql = @" update OQCTestProgSet set sampleplan ='" + sampleplan + "', MA ='" + MA + "', MAvalue=" + MAvalue + ", MI ='" + MI + "', MIvalue =" + MIvalue + ",IPCA610F = '" + IPCA610F + "',customerstandard ='" + customerstandard + "', otherstandard = '" + otherstandard + "',SNstates = '"+ SNstates + "' where customer ='' and PN =''";
                bool flag = DbAccess.ExecuteSql(sql);
                if (flag)
                {
                    MessageBox.Show("修改成功！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    sBtnadd_Click(sender, e);
                }
                else
                {
                    MessageBox.Show("修改失败！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("不能只对编码修改！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
               
            }

            /*
            string flag = OQCTestProgSetAdd("修改", customer, PN, model, testitem, checkmethod, int.Parse(standardsequence), teststandard, checkMA, checkMI, sampleplan, MA,MAvalue, MI, MIvalue, IPCA610F, customerstandard, otherstandard, "");
            if (flag == "修改成功")
            {
                MessageBox.Show(flag, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                sBtnadd_Click(sender, e);
            }
            else
                MessageBox.Show(flag, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
             */
        }



        DataTable dt = new DataTable();
        string connString = "server=192.168.0.204;database=BarcodeNew;user id=sa;Pooling=false;password=The0more7people0you7love3the7weaker8you8are";
        SqlConnection conn;
        private void bind(string fileName)
        {
            string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" +
                 "Data Source=" + fileName + ";" +
                 "Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'";
            //  string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
            OleDbConnection conn = new OleDbConnection(strConn);
            if (conn.State == ConnectionState.Broken || conn.State == ConnectionState.Closed)
            {
                conn.Open();
            }
             System.Data.DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            string sheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString().Trim();

            string strSql = "select * from [" + sheetName + "]";
            OleDbDataAdapter da = new OleDbDataAdapter(strSql, conn);
            // OleDbDataAdapter da = new OleDbDataAdapter("SELECT *  FROM [Sheet1$]", strConn);
            DataSet ds = new DataSet();
            try
            {
                da.Fill(ds);
                dt = ds.Tables[0];
                this.gridControltosql.DataSource = dt;
            }
            catch (Exception err)
            {
                MessageBox.Show("操作失败！" + err.ToString());
            }
        }




        public DataTable CreateTestItemTable()
        {

            DataTable namesTable = new DataTable("TestItem");
            DataColumn testitem = new DataColumn();
            testitem.DataType = System.Type.GetType("System.String");
            testitem.ColumnName = "testitem";
            namesTable.Columns.Add(testitem);

            DataColumn checkmethod = new DataColumn();
            checkmethod.DataType = System.Type.GetType("System.String");
            checkmethod.ColumnName = "checkmethod";
            namesTable.Columns.Add(checkmethod);

            DataColumn standardsequence = new DataColumn();
            standardsequence.DataType = System.Type.GetType("System.Int32");
            standardsequence.ColumnName = "standardsequence";
            namesTable.Columns.Add(standardsequence);

            DataColumn teststandard = new DataColumn();
            teststandard.DataType = System.Type.GetType("System.String");
            teststandard.ColumnName = "teststandard";
            namesTable.Columns.Add(teststandard);

            DataColumn checkMA = new DataColumn();
            checkMA.DataType = System.Type.GetType("System.String");
            checkMA.ColumnName = "checkMA";
            namesTable.Columns.Add(checkMA);

            DataColumn checkMI = new DataColumn();
            checkMI.DataType = System.Type.GetType("System.String");
            checkMI.ColumnName = "checkMI";
            namesTable.Columns.Add(checkMI);

            return namesTable;

        }

      
        private Microsoft.Office.Interop.Excel.ApplicationClass appClsExcel = null;
        Microsoft.Office.Interop.Excel.Workbooks wbs = null;
        Microsoft.Office.Interop.Excel.Workbook wb = null;
        Microsoft.Office.Interop.Excel.Worksheet ws = null;
        Microsoft.Office.Interop.Excel.Range range = null;
        private void OpenExcel(string strFileName)
        {
            object objMissing = System.Reflection.Missing.Value;
            this.appClsExcel = new Microsoft.Office.Interop.Excel.ApplicationClass();
            appClsExcel.UserControl = true;
            appClsExcel.DisplayAlerts = false;

            Microsoft.Office.Interop.Excel.Workbook wBookExcel = appClsExcel.Workbooks.Open(strFileName, objMissing, true, objMissing, objMissing, objMissing
                          , objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing, objMissing);

            Microsoft.Office.Interop.Excel.Worksheet wSheetExcel = (Microsoft.Office.Interop.Excel.Worksheet)wBookExcel.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range rCell = wSheetExcel.UsedRange;

            object[,] objList = (object[,])rCell.Value2;

            int excelcounts = objList.GetLength(0);

            wbs = appClsExcel.Workbooks;
            wb = wbs[1];
            ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets["Sheet1"];
            int rowCount = ws.UsedRange.Rows.Count;
            int colCount = ws.UsedRange.Columns.Count;
            if (rowCount <= 0)
            {
                MessageBox.Show("文件中没有数据记录", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (colCount < 6)
            {
                MessageBox.Show("字段个数不对", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DataTable TestItemTable = CreateTestItemTable();

            //Microsoft.Office.Interop.Excel.Range rng1 = ws.Cells.get_Range("B2", "B" + rowCount);
            //Microsoft.Office.Interop.Excel.Range rng2 = ws.Cells.get_Range("C2", "C" + rowCount);
            //Microsoft.Office.Interop.Excel.Range rng3 = ws.Cells.get_Range("D2", "D" + rowCount);
            //Microsoft.Office.Interop.Excel.Range rng4 = ws.Cells.get_Range("E2", "E" + rowCount);
            //Microsoft.Office.Interop.Excel.Range rng5 = ws.Cells.get_Range("F2", "F" + rowCount);
            //Microsoft.Office.Interop.Excel.Range rng6 = ws.Cells.get_Range("G2", "G" + rowCount);

            //object[,] testitem = (object[,])rng1.Value2;  
            //object[,] checkmethod = (object[,])rng2.Value2;
            //object[,] standardsequence = (object[,])rng3.Value2;
            //object[,] teststandard = (object[,])rng4.Value2;
            //object[,] checkMA = (object[,])rng5.Value2;
            //object[,] checkMI = (object[,])rng6.Value2;

            //string testitem1 = "", checkmethod1 = "", standardsequence1 = "", teststandard1 = "", checkMA1 = "", checkMI1 = "";

            for (int iRow = 1; iRow < rowCount; iRow++)
                {
                //testitem1 = testitem[iRow, 1].ToString();
                //checkmethod1 = checkmethod[iRow, 1].ToString();
                //standardsequence1 = standardsequence[iRow, 1].ToString();
                //teststandard1 = teststandard[iRow, 1].ToString();
                //checkMA1 = checkMA[iRow, 1].ToString();
                //checkMI1 = checkMI[iRow, 1].ToString();

                //if (testitem1 != "" && checkmethod1 != "" && standardsequence1 != "" && teststandard1 != "" && checkMA1 != "" && checkMI1 != "")
                //{
                //    DataRow row;
                //    row = TestItemTable.NewRow();
                //    row["testitem"] = testitem1;
                //    row["checkmethod"] = checkmethod1;
                //    row["standardsequence"] = standardsequence1;
                //    row["teststandard"] = teststandard1;
                //    row["checkMA"] = checkMA1;
                //    row["checkMI"] = checkMI1;
                //    TestItemTable.Rows.Add(row);
                //}
                //else
                //{
                //    break;
                //}
                if (objList[iRow + 1, 2].ToString() != "" && objList[iRow + 1, 3].ToString() != "" && objList[iRow + 1, 4].ToString() != "" && objList[iRow + 1, 5].ToString() != "" && objList[iRow + 1, 6].ToString() != "" && objList[iRow + 1, 7].ToString() != "")
                {
                    DataRow row;
                    row = TestItemTable.NewRow();
                    row["testitem"] = objList[iRow + 1, 2].ToString();
                    row["checkmethod"] = objList[iRow + 1, 3].ToString();
                    row["standardsequence"] = objList[iRow + 1, 4].ToString();
                    row["teststandard"] = objList[iRow + 1, 5].ToString();
                    row["checkMA"] = objList[iRow + 1, 6].ToString();
                    row["checkMI"] = objList[iRow + 1, 7].ToString();
                    TestItemTable.Rows.Add(row);
                }
                else
                {
                    break;
                }

            }
            this.gridControltosql.DataSource = TestItemTable;

            appClsExcel.Quit();
            appClsExcel = null;
            Process[] procs = Process.GetProcessesByName("Excel");
            foreach (Process pro in procs)
            {
                pro.Kill();
            }
            GC.Collect();
        }



        private void sBtninsert_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog fd = new OpenFileDialog();
            if (fd.ShowDialog() == DialogResult.OK)
            {
                string fileExtenSion;
                fileExtenSion = Path.GetExtension(fd.FileName);
                if (fileExtenSion.ToLower() != ".xls" && fileExtenSion.ToLower() != ".xlsx")
                {
                    MessageBox.Show("文件格式不正确", "停止",MessageBoxButtons.OK,MessageBoxIcon.Stop);
                    return ;
                }
                // bind(fd.FileName);

                try
                {
                    OpenExcel(fd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导入失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

            }

        }  
        private void insertToSql(DataRow dr)
        {
            string customer = dr["customer"].ToString().Trim ();
            string PN = dr["PN"].ToString().Trim ();
            string testitem = dr["testitem"].ToString().Trim();
            string checkmethod = dr["checkmethod"].ToString().Trim ();
            string standardsequence = dr["standardsequence"].ToString();
            string teststandard = dr["teststandard"].ToString();
            string checkMA = dr["checkMA"].ToString();
            string checkMI = dr["checkMI"].ToString();
            string flag =  OQCTestItemAdd("新增",customer,PN,testitem, checkmethod, int.Parse(standardsequence), teststandard, checkMA, checkMI);
        }  
        private void sbtntijiao_Click(object sender, EventArgs e)
        {
            /*
            conn = new SqlConnection(connString);
            conn.Open();
            DataTable dtt = this.gridControltosql.DataSource as DataTable;
            if (dtt.Rows.Count > 0)
            {
                DataRow dr = null;
                for (int i = 0; i < dtt.Rows.Count; i++)
                {
                    dr = dtt.Rows[i];
                    insertToSql(dr);
                }
                conn.Close();
                MessageBox.Show("导入成功！");
            }
            else
            {
                MessageBox.Show("没有数据！");
            }
            */
        }

        private void switchedit_Toggled(object sender, EventArgs e)
        {
            if (switchedit.IsOn)
            {
                gridViewtosql.OptionsBehavior.Editable = true;
                sBtnsubmit.Enabled = false;
            }
            else
            {
                gridViewtosql.OptionsBehavior.Editable = false;
                sBtnsubmit.Enabled = true;
            }
        }

        private void bandedGridView_Click(object sender, EventArgs e)
        {
            if (bandedGridView.RowCount < 0)
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DataTable de = gridControl.DataSource as DataTable;
            if (de == null)
            {
                MessageBox.Show("没有数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            cboBcustomer.Text = bandedGridView.GetFocusedRowCellValue("客户").ToString().Trim();
            txtPN.Text = bandedGridView.GetFocusedRowCellValue("编码").ToString().Trim();
            txtmodel.Text = bandedGridView.GetFocusedRowCellValue("机型").ToString().Trim();
        }

        private void sBtnExport_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 0)
                return;
            int ii = bandedGridView.FocusedRowHandle;
            if (ii < 0)
            {
                MessageBox.Show("请选中需要导出的项目", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            string customer = bandedGridView.GetFocusedRowCellValue("客户").ToString();
            string PN = bandedGridView.GetFocusedRowCellValue("编码").ToString();
            string model = bandedGridView.GetFocusedRowCellValue("机型").ToString();
            string sql = "";
            if (customer != "" && PN != "")
            {
                sql = @"  select testitem,checkmethod,standardsequence,teststandard,checkMA,checkMI from OQCTestProgSet where customer ='"+customer+ "' and PN='"+PN+ "' order by testitem ,standardsequence ";
            }
            else if (customer != "" && PN == "")
            {
                sql = @"  select testitem,checkmethod,standardsequence,teststandard,checkMA,checkMI from OQCTestProgSet where customer ='" + customer + "' and PN='' order by testitem ,standardsequence ";
            }
            else if (customer == "" && PN == "")
            {
                sql = @"  select testitem,checkmethod,standardsequence,teststandard,checkMA,checkMI from OQCTestProgSet where customer ='' and PN='' order by testitem ,standardsequence ";
            }
            else
            {
                MessageBox.Show("不能只对编码导出", "提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }

            DataTable dtt = DbAccess.SelectBySql(sql).Tables[0];
            if (dtt == null || dtt.Rows.Count < 0)
                return;


            try
            {
                int sheetCount = 1;
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = true;
                object missing = System.Reflection.Missing.Value;
                string templetFile = Environment.CurrentDirectory + @"\ReportFolder\OQC测试项目.xlsx";
                Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(templetFile, missing, true, missing, missing, missing,
                                                              missing, missing, missing, missing, missing, missing, missing, missing, missing);
                Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets.get_Item(1);

                for (int i = 1; i < sheetCount; i++)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(i)).Copy(missing, workBook.Worksheets[i]);
                }
                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);
                if (sheet == null)
                    return;


                for (int n = 0, m = 2; n < dtt.Rows.Count ; n++, m++)
                {
                    sheet.Cells[m, 2] = dtt.Rows[n]["testitem"].ToString();
                    sheet.Cells[m, 3] = dtt.Rows[n]["checkmethod"].ToString();
                    sheet.Cells[m, 4] = dtt.Rows[n]["standardsequence"].ToString();
                    sheet.Cells[m, 5] = dtt.Rows[n]["teststandard"].ToString();
                    sheet.Cells[m, 6] = dtt.Rows[n]["checkMA"].ToString();
                    sheet.Cells[m, 7] = dtt.Rows[n]["checkMI"].ToString();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("导出失败！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}