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
using System.Data.SqlClient;
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Base;
using System.IO;

namespace DX_QMS
{
    public partial class ReturnSearsh : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        private string serverTinFile = "QMSSVR\\TinReport";
        IQC ic = new IQC();
        DataTable dt = null;
        public ReturnSearsh()
        {
            InitializeComponent();
            bindDeviceType();
            bindTestItem(txttesttype.SelectedItem.ToString());
            bindTestTools();
        }
        private void bindDeviceType()
        {
            System.Data.DataTable dt = ic.SelectTestTypeRecord("查询", "", "测试类别", "").Tables[0];
            txttesttype.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txttesttype.Properties.Items.Add(row["TestType"]);
            }
            txttesttype.SelectedIndex = 0;
        }
        private void bindTestItem(string TestType)
        {
            DataSet ds = ic.SelectTestItemRecord("查询", TestType, "", 0, "");
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                txttestitem.Properties.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    txttestitem.Properties.Items.Add(row["TestItem"]);
                }
                    txttestitem.SelectedIndex = 0;  
            }
        }

        private void bindTestTools()
        {
            System.Data.DataTable dt = ic.SelectTestTypeRecord("查询", "", "测试工具", "").Tables[0];
            txttesttools.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                txttesttools.Properties.Items.Add(row["TestType"]);
            }
        }

        private void txttesttype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txttesttype.Text.ToString() == "")
                return;
            bindTestItem(txttesttype.Text.ToString());
        }
        private void bindTestSubItem(string testtype, string Testitem)
        {
            //string ssql = "select * from IQC_TestSubItemSet where TestType=case " + testtype + " when '' then TestType else " + testtype + " end and TestItem= case " + Testitem + " when '' then TestItem else " + Testitem + " end";
            //DataSet ds = DbAccess.SelectBySql(ssql);
            //if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            //{
            //    System.Data.DataTable dt = ds.Tables[0];
            //    txttestsubitem.Properties.Items.Clear();
            //    foreach (DataRow row in dt.Rows)
            //    {
            //        txttestsubitem.Properties.Items.Add(row["TestSubItem"]);
            //    }
            //        txttestsubitem.SelectedIndex = 0;
            //}

            string ssql = "select * from IQC_TestSubItemSet where TestType=case '" + testtype + "' when '' then TestType else '" + testtype + "' end and TestItem= case '" + Testitem + "' when '' then TestItem else '" + Testitem + "' end";
            DataSet ds = Common.DbAccess.SelectBySql(ssql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                System.Data.DataTable dt = ds.Tables[0];
                txttestsubitem.Properties.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    txttestsubitem.Properties.Items.Add(row["TestSubItem"]);
                }
                txttestsubitem.SelectedIndex = 0;
            }


        }
        private void txttestitem_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txttestitem.Text.ToString() == "")
                return;
                //bindTestSubItem(txttesttype.SelectedItem.ToString() == null ? "" : txttesttype.SelectedItem.ToString(), txttestitem.SelectedItem.ToString() == null ? "" : txttestitem.SelectedItem.ToString());
                bindTestSubItem(txttesttype.Text == "" ? "" : txttesttype.Text, txttestitem.Text == "" ? "" : txttestitem.Text);
        }

        private void ReturnSearsh_Load(object sender, EventArgs e)
        {
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }


        public DataSet SelectTestListRecord(string opertype, string testtype, string testItem, string TestSubItem, string TestTool, string lotno, string userid, string materialcode, string materialname, string vendorcode, string vendorname, string bdate, string edate)
        {
            SqlParameter[] para = new SqlParameter[13];
            para[0] = new SqlParameter("@opertype", opertype);
            para[1] = new SqlParameter("@testtype", testtype);
            para[2] = new SqlParameter("@testitem", testItem);
            para[3] = new SqlParameter("@testsubitem", TestSubItem);
            para[4] = new SqlParameter("@testtools", TestTool);
            para[5] = new SqlParameter("@lotno", lotno);
            para[6] = new SqlParameter("@testuser", userid);
            para[7] = new SqlParameter("@materialcode", materialcode);
            para[8] = new SqlParameter("@materialname", materialname);
            para[9] = new SqlParameter("@vendorcode", vendorcode);
            para[10] = new SqlParameter("@vendorname", vendorname);
            para[11] = new SqlParameter("@testtime1", bdate);
            para[12] = new SqlParameter("@testtime2", edate);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "IQC_TestListSearchNew", para);
        }

        private void sipBtnsearsh_Click(object sender, EventArgs e)
        {
            //DataSet ds = null;
            //ds = SelectTestListRecord("退料测试详细", txttesttype.SelectedItem == null ? "" : txttesttype.SelectedItem.ToString(), txttestitem.SelectedItem == null ? "" : txttestitem.SelectedItem.ToString(), txttestsubitem.SelectedItem == null ? "" : txttestsubitem.SelectedText, txttesttools.SelectedItem == null ? "" : txttesttools.SelectedItem.ToString(), txtlotno.Text, txttestuser.Text, txtmaterialcode.Text, txtmaterialname.Text, txtvendorcode.Text, txtvendorname.Text, bdate.Text, edate.Text);
            //if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            //{
            //    gridControl.DataSource = ds.Tables[0];
            //}
            //else
            //{
            //    MessageBox.Show("没有符合条件的记录");
            //    gridControl.DataSource = null;
            //}


            string sqlstr = "", sqlwhere1 = " " ,sqlwhere2 = " where 1=1";
            string materialcode = "", materialname = "", vendorcode = "", vendorname = "", returntype = "";
            string testtype = "", testitem = "", testsubitem = "", testtools = "", testuser = "", lotno = "",deliveryid = "";

            materialcode = txtmaterialcode.Text.Trim();
            materialname = txtmaterialname.Text.Trim();
            vendorcode = txtvendorcode.Text.Trim();
            vendorname = txtvendorname.Text.Trim();
            returntype = returncategory.Text.Trim();

            testtype = txttesttype.Text.Trim();
            testitem = txttestitem.Text.Trim();
            testsubitem = txttestsubitem.Text.Trim();
            testtools = txttesttools.Text.Trim();
            testuser = txttestuser.Text.Trim();
            lotno = txtlotno.Text.Trim();
            deliveryid = txtdeliveryid.Text.Trim();


            string sbdate = "", sedate = "";
            sbdate = bdate.Text;
            sedate = edate.Text;

            gridControl.DataSource = null;
            DataSet ds = null;

            //if (!string.IsNullOrEmpty(materialcode))
            //{
            //    sqlwhere1 += " and d.materialcode like '%" + materialcode + "%' ";
            //}
            //if (!string.IsNullOrEmpty(deliveryid))
            //{
            //    sqlwhere1 += " and d.deliveryid = '" + deliveryid + "' ";
            //}
            //if (!string.IsNullOrEmpty(materialname))
            //{
            //    sqlwhere1 += "  and materialname like '%" + materialname + "%' ";
            //}
            //if (!string.IsNullOrEmpty(vendorcode))
            //{
            //    sqlwhere1 += " and vendorcode like '%" + vendorcode + "%' ";
            //}
            //if (!string.IsNullOrEmpty(vendorname))
            //{
            //    sqlwhere1 += "  and vendorname like '%" + vendorname + "%' ";
            //}

            //if (!string.IsNullOrEmpty(testtype))
            //{
            //    sqlwhere2 += " and testtype='" + testtype + "' ";
            //}
            //if (!string.IsNullOrEmpty(testitem))
            //{
            //    sqlwhere2 += " and testitem='" + testitem + "' ";
            //}
            //if (!string.IsNullOrEmpty(testsubitem))
            //{
            //    sqlwhere2 += " and testsubitem='" + testsubitem + "' ";
            //}
            //if (!string.IsNullOrEmpty(testtools))
            //{
            //    sqlwhere2 += "  and TestTools like '%" + testtools + "%' ";
            //}
            //if (!string.IsNullOrEmpty(returntype))
            //{
            //    sqlwhere2 += "  and returnreason like '%" + returntype + "%' ";
            //}
            //if (!string.IsNullOrEmpty(testuser))
            //{
            //    sqlwhere2 += " and TestUser like '%" + testuser + "%' ";
            //}
            //if (!string.IsNullOrEmpty(lotno))
            //{
            //    sqlwhere2 += "  and Lotno like '%" + lotno + "%' ";
            //}
            //if (!string.IsNullOrEmpty(sbdate))
            //{
            //    sbdate = bdate.DateTime.ToString("yyyy-MM-dd");
            //   sqlwhere2 += " and testtime >= '" + sbdate + " 00:00:00 ' ";
            //}
            //if (!string.IsNullOrEmpty(sedate))
            //{
            //    sedate = edate.DateTime.ToString("yyyy-MM-dd");
            //    sqlwhere2 += " and testtime <= '" + sedate + " 23:59:59 ' ";
            //}

            //sqlstr = "  select d.deliveryid 接收单号,d.materialcode 料号,max(materialname) 名称,case when max(returnqty)>1 then max(returnqty) else max(x.Qty) end  退料数量,max(x.returnreason) 退料原因,max(x.Badcategory) 不良类别,max(vendorname) 供应商名称,min(transactiondate) 生成时间, ";
            //    sqlstr += " case max(TestFinalResult) when '拒收' then 'NG' else 'OK' end 检验结果,max(testtime) 测试日期,max(TestType) 测试类别,Max(Remarks) 备注,max(TestUser) 检验员  from delivery d left join MaterialSpec m on d.materialcode=m.materialcode  ";
            //    sqlstr += "  right join  (  select receptid,Productcode,LotNo,max(returnqty) returnqty,max(Qty) Qty,max(TestFinalResult) TestFinalResult,max(returnreason) returnreason,max(Badcategory) Badcategory,min(testtime) testtime,max(TestType) TestType,MAX(Remarks) Remarks,max(TestUser) TestUser from IQC_TestListReturn ";
            //    sqlstr += "  "+sqlwhere2+ " and returnqty is not null group by receptid,Productcode,Lotno  ";
            //    sqlstr += "  ) x on d.deliveryid=x.receptid and d.materialcode=x.Productcode and d.lotno=x.LotNo   ";
            //    sqlstr += "  "+ sqlwhere1 + "  group by d.deliveryid,d.materialcode  order by 测试日期 desc  ";

           
                if (!string.IsNullOrEmpty(materialcode))
                {
                    sqlwhere1 += " and r.Productcode like '%" + materialcode + "%' ";
                }
                if (!string.IsNullOrEmpty(deliveryid))
                {
                    sqlwhere1 += " and r.receptid = '" + deliveryid + "' ";
                }
                if (!string.IsNullOrEmpty(materialname))
                {
                    sqlwhere1 += "  and m.materialname like '%" + materialname + "%' ";
                }
                if (!string.IsNullOrEmpty(vendorcode))
                {
                    sqlwhere1 += " and d.vendorcode like '%" + vendorcode + "%' ";
                }
                if (!string.IsNullOrEmpty(vendorname))
                {
                    sqlwhere1 += "  and d.vendorname like '%" + vendorname + "%' ";
                }
                if (!string.IsNullOrEmpty(returntype))
                {
                    sqlwhere1 += "  and r.returnreason = '" + returntype + "' ";
                }

                if (!string.IsNullOrEmpty(testtype))
                {
                    sqlwhere1 += " and r.TestType='" + testtype + "' ";
                }
                if (!string.IsNullOrEmpty(testitem))
                {
                    sqlwhere1 += " and r.TestItem='" + testitem + "' ";
                }
                if (!string.IsNullOrEmpty(testsubitem))
                {
                    sqlwhere1 += " and r.TestSubItem='" + testsubitem + "' ";
                }
                if (!string.IsNullOrEmpty(testtools))
                {
                    sqlwhere1 += "  and r.TestTools like '%" + testtools + "%' ";
                }
                if (!string.IsNullOrEmpty(testuser))
                {
                    sqlwhere1 += " and r.TestUser like '%" + testuser + "%' ";
                }
                if (!string.IsNullOrEmpty(lotno))
                {
                    sqlwhere1 += "  and r.LotNo like '%" + lotno + "%' ";
                }
                if (!string.IsNullOrEmpty(sbdate))
                {
                    sbdate = bdate.DateTime.ToString("yyyy-MM-dd");
                    sqlwhere1 += " and r.TestTime >= '" + sbdate + " 00:00:00 ' ";
                }
                if (!string.IsNullOrEmpty(sedate))
                {
                    sedate = edate.DateTime.ToString("yyyy-MM-dd");
                    sqlwhere1 += " and r.TestTime <= '" + sedate + " 23:59:59 ' ";
                }

            if (cbExpiryDate.Checked == false)
            {
                if (sqlwhere1.Trim() != "")
                {
                    sqlstr = "  select receptid 接收单号,Productcode 料号,r.LotNo 批次号,'' 抽检流水号,max(materialname) 名称,max(d.vendorname) 供应商名称,max(returnqty) 重检数量,max(r.returnreason) 退料原因,max(r.Badcategory) 不良类别,case max(TestFinalResult) when '拒收' then 'NG' else 'OK' end 检验结果,min(testtime) 测试日期,max(TestType) 测试类别,MAX(Remarks) 备注,max(TestUser) 检验员  ";
                    sqlstr += " from IQC_TestListReturn r left join delivery d on d.deliveryid=r.receptid and d.materialcode=r.Productcode and d.lotno=r.LotNo left join   MaterialSpec m on r.Productcode = m.materialcode  ";
                    sqlstr += "   where  returnqty is not null " + sqlwhere1;
                    sqlstr += "   group by receptid,r.Productcode,r.Lotno order by 测试日期 desc    ";
                }
                else
                {
                    sqlstr = "  select  top 1000  receptid 接收单号,Productcode 料号,r.LotNo 批次号,'' 抽检流水号,  max(materialname) 名称,max(d.vendorname) 供应商名称,max(returnqty) 重检数量,max(r.returnreason) 退料原因,max(r.Badcategory) 不良类别,case max(TestFinalResult) when '拒收' then 'NG' else 'OK' end 检验结果,min(testtime) 测试日期,max(TestType) 测试类别,MAX(Remarks) 备注,max(TestUser) 检验员  ";
                    sqlstr += " from IQC_TestListReturn r left join delivery d on d.deliveryid=r.receptid and d.materialcode=r.Productcode and d.lotno=r.LotNo left join   MaterialSpec m on r.Productcode = m.materialcode  ";
                    sqlstr += "   where  returnqty is not null " + sqlwhere1;
                    sqlstr += "   group by receptid,r.Productcode,r.Lotno order by 测试日期 desc    ";
                }
            }
            else
            {
                if (sqlwhere1.Trim() != "")
                {
                    sqlstr = "  select receptid 接收单号,Productcode 料号,r.LotNoList 批次号,r.CheckItem 抽检流水号,max(materialname) 名称,max(d.vendorname) 供应商名称,max(returnqty) 重检数量,max(r.returnreason) 退料原因,max(r.Badcategory) 不良类别,case max(TestFinalResult) when '拒收' then 'NG' else 'OK' end 检验结果,min(testtime) 测试日期,max(TestType) 测试类别,MAX(Remarks) 备注,max(TestUser) 检验员  ";
                    sqlstr += " from IQC_TestListReturn r left join delivery d on d.deliveryid=r.receptid and d.materialcode=r.Productcode left join   MaterialSpec m on r.Productcode = m.materialcode  ";
                    sqlstr += "   where  returnqty is not null and LotNoList is not null " + sqlwhere1;
                    sqlstr += "   group by receptid,r.Productcode,LotNoList,CheckItem order by 测试日期 desc    ";
                }
                else
                {
                    sqlstr = "  select  top 1000  receptid 接收单号,Productcode 料号,r.LotNoList 批次号,r.CheckItem 抽检流水号,  max(materialname) 名称,max(d.vendorname) 供应商名称,max(returnqty) 重检数量,max(r.returnreason) 退料原因,max(r.Badcategory) 不良类别,case max(TestFinalResult) when '拒收' then 'NG' else 'OK' end 检验结果,min(testtime) 测试日期,max(TestType) 测试类别,MAX(Remarks) 备注,max(TestUser) 检验员  ";
                    sqlstr += " from IQC_TestListReturn r left join delivery d on d.deliveryid=r.receptid and d.materialcode=r.Productcode left join   MaterialSpec m on r.Productcode = m.materialcode  ";
                    sqlstr += "   where  returnqty is not null and LotNoList is not null " + sqlwhere1;
                    sqlstr += "   group by receptid,r.Productcode,LotNoList,CheckItem order by 测试日期 desc    ";
                }
            }
             DataTable dt = DbAccess.SelectBySql(sqlstr).Tables[0];

            if (dt != null && dt.Rows.Count > 0)
            {
                gridControl.DataSource = dt;
            }
            else
            {
                MessageBox.Show("没有符合条件的记录");
                gridControl.DataSource = null;
            }

        }

        private void sipBtnreset_Click(object sender, EventArgs e)
        {
            txttesttype.Text = "";
            txttestitem.Text = "";
            txttestsubitem.Text = "";
            txttesttools.Text = "";
            txttestuser.Text = "";
            returncategory.Text = "";
            txtmaterialname.Text = "";
            txtvendorcode.Text = "";
            txtvendorname.Text = "";
            txtlotno.Text = "";
            txtmaterialcode.Text = "";
            txtdeliveryid.Text = "";
            bdate.Text = "";
            edate.Text = "";
        }
        private string GetReAc(string Sampleqty, string AQL)
        {
            string AR = "";
            string ssql = "select Sampleqty,AQL,cast(Ac as varchar(10))+'/'+cast(Re as varchar(10)) AR from IQC_TestSTD105ECode s left join  IQC_TestAQLRcSet r on s.Code=r.Code where Sampleqty='" + Sampleqty + "' and AQL='" + AQL + "'";
            DataSet ds = DbAccess.SelectBySql(ssql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                AR = ds.Tables[0].Rows[0]["AR"].ToString();
            }
            return AR;
        }

        public DateTime beforeTime, afterTime;
        public void Copy(string sheetPrefixName, string sup, string dcode, string orderno, System.Data.DataTable tb)
        {
            string finaltotalstate = "", testuser = "";
            int sheetCount = 1;

            if (sheetPrefixName == null || sheetPrefixName.Trim() == "")
                sheetPrefixName = " Sheet ";
            beforeTime = DateTime.Now;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            afterTime = DateTime.Now;
            object missing = System.Reflection.Missing.Value;
            string templetFile = Environment.CurrentDirectory + @"\Resources\Report.xls";
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(templetFile, missing, true, missing, missing, missing,
                                                          missing, missing, missing, missing, missing, missing, missing, missing, missing);
            Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets.get_Item(1);
            for (int i = 1; i < sheetCount; i++)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(i)).Copy(missing, workBook.Worksheets[i]);
            }
            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);
            sheet.Name = sheetPrefixName.Replace("/", "");

            string lotnos = "";
            System.Data.DataTable dtlotno = null;
            for (int j = 0; j < tb.Rows.Count; j++)
            {
                string finalstate = "";
                sheet.Cells[2, 1] = "IQC检验报表(" + dt.Rows[0 + j]["TestType"].ToString() + ")";
                sheet.Cells[4, 2] = sup;
                sheet.Cells[4, 18] = dcode;
                //sheet.Cells[4, 7] = orderno;
                sheet.Cells[3, 2] = dt.Rows[0 + j]["TestTime"].ToString();
                sheet.Cells[3, 6] = dt.Rows[0 + j]["lot_number"].ToString();
                sheet.Cells[4, 12] = dt.Rows[0 + j]["Qty"].ToString();
                sheet.Cells[3, 17] = dt.Rows[0 + j]["LotNo"].ToString();
                sheet.Cells[46, 19] = dt.Rows[0 + j]["TestUser"].ToString();
                sheet.Cells[5, 2] = dt.Rows[0 + j]["Productcode"].ToString();
                sheet.Cells[5, 7] = dt.Rows[0 + j]["materialenname"].ToString();
                sheet.Cells[10 + j + 1, 1] = dt.Rows[0 + j]["TestItem"].ToString();
                sheet.Cells[10 + j + 1, 2] = dt.Rows[0 + j]["TestSubItem"].ToString();
                sheet.Cells[10 + j + 1, 3] = dt.Rows[0 + j]["TestDes"].ToString();
                sheet.Cells[10 + j + 1, 9] = dt.Rows[0 + j]["TestTools"].ToString();
                sheet.Cells[10 + j + 1, 11] = dt.Rows[0 + j]["SampleType"].ToString();
                if (dt.Rows[0 + j]["AQL"].ToString().StartsWith("MI"))
                {
                    sheet.Cells[10 + j + 1, 20] = "√";

                }
                else if (dt.Rows[0 + j]["AQL"].ToString().StartsWith("MA"))
                {
                    sheet.Cells[10 + j + 1, 19] = "√";
                }
                else if (dt.Rows[0 + j]["AQL"].ToString().StartsWith("CR=0.1"))
                {
                    sheet.Cells[10 + j + 1, 18] = "√";
                }

                else if (dt.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.40"))
                {
                    sheet.Cells[10 + j + 1, 21] = "√";
                }
                else if (dt.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.0"))
                {
                    sheet.Cells[10 + j + 1, 22] = "√";
                }
                else if (dt.Rows[0 + j]["AQL"].ToString().StartsWith("CR=0.01"))
                {
                    sheet.Cells[10 + j + 1, 23] = "√";
                }

                sheet.Cells[10 + j + 1, 17] = GetReAc(dt.Rows[0 + j]["Sampleqty"].ToString(), dt.Rows[0 + j]["AQL"].ToString());

                if (dt.Rows[0 + j]["CheckType"].ToString() == "正常检验")
                {
                    // sheet.Cells[7, 6] = "√正常检验";
                    sheet.Cells.get_Range("F7").Value = "√正常检验";
                    sheet.Cells.get_Range("F7").Font.Size = 11;
                    sheet.Cells.get_Range("F7").Font.Bold = true;
                    sheet.Cells.get_Range("F7").EntireColumn.AutoFit();
                    sheet.Cells.get_Range("F7").EntireRow.AutoFit();
                    sheet.Cells.get_Range("F7").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LightGreen);
                }
                else if (dt.Rows[0 + j]["CheckType"].ToString() == "加严检验")
                {
                    // sheet.Cells[7, 2] = "√加严检验";
                    sheet.Cells.get_Range("B7").Value = "√加严检验";
                    sheet.Cells.get_Range("B7").Font.Size = 11;
                    sheet.Cells.get_Range("B7").Font.Bold = true;
                    sheet.Cells.get_Range("B7").EntireColumn.AutoFit();
                    sheet.Cells.get_Range("B7").EntireRow.AutoFit();
                    sheet.Cells.get_Range("B7").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);

                }
                else if (dt.Rows[0 + j]["CheckType"].ToString() == "放宽检验")
                {
                    // sheet.Cells[7, 11] = "√放宽检验";
                    sheet.Cells.get_Range("K7").Value = "√放宽检验";
                    sheet.Cells.get_Range("K7").Font.Size = 11;
                    sheet.Cells.get_Range("K7").Font.Bold = true;
                    sheet.Cells.get_Range("K7").EntireColumn.AutoFit();
                    sheet.Cells.get_Range("K7").EntireRow.AutoFit();
                    sheet.Cells.get_Range("K7").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LightYellow);

                }
                else if (dt.Rows[0 + j]["CheckType"].ToString() == "全检")
                {
                    //sheet.Cells[7, 16] = "√全检";
                    sheet.Cells.get_Range("P7").Value = "√全检";
                    sheet.Cells.get_Range("P7").Font.Size = 11;
                    sheet.Cells.get_Range("P7").Font.Bold = true;
                    sheet.Cells.get_Range("P7").EntireColumn.AutoFit();
                    sheet.Cells.get_Range("P7").EntireRow.AutoFit();
                    //sheet.Cells.get_Range("P7").Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);
                }

                System.Data.DataTable dlist = ic.SelectTestListRecord("测试明细", "", tb.Rows[j]["TestItem"].ToString(), tb.Rows[j]["TestSubItem"].ToString(), "", tb.Rows[j]["lotno"].ToString(), "", tb.Rows[j]["Productcode"].ToString(), "", "", "", "", "").Tables[0];
                dtlotno = dlist;
                string s = "";
                int testqty = 0;
                for (int m = 0; m < dlist.Rows.Count; m++)
                {
                    testqty = testqty + 1;
                    s = s + (m + 1).ToString() + ".测试值:" + dlist.Rows[0 + m]["Testvalue"].ToString() + ",检验结果:" + dlist.Rows[0 + m]["TestResult"].ToString() + ",备注:" + dlist.Rows[0 + m]["Remarks"].ToString() + ";";
                    finalstate = finalstate + dlist.Rows[0 + m]["TestFinalResult"].ToString();
                }

                finaltotalstate = finaltotalstate + finalstate;
                testuser = tb.Rows[j]["TestUser"].ToString();

                //sheet.Cells[14 + j + 1, 11] = testqty;
                sheet.Cells[10 + j + 1, 12] = s.TrimEnd(';').TrimStart(';');
                if (finalstate.Contains("拒收"))
                    //sheet.Cells[10 + j + 1, 21] = "NG";
                    sheet.Cells[10 + j + 1, 24] = "NG";
                else
                    //sheet.Cells[10 + j + 1, 21] = "OK";
                    sheet.Cells[10 + j + 1, 24] = "OK";
            }

            //20141030新增加的

            System.Data.DataTable newtb = dtlotno.Clone();
            newtb.Rows.Add(dtlotno.Rows[0].ItemArray);
            lotnos = lotnos + "," + dtlotno.Rows[0]["LotNos"].ToString();
            for (int i = 1; i < dtlotno.Rows.Count; i++)
            {
                bool flag = true;
                foreach (DataRow dr in newtb.Rows)
                {
                    if (dtlotno.Rows[i]["LotNos"].ToString() == dr["LotNos"].ToString())
                    {
                        flag = false;
                        continue;
                    }
                }
                if (flag)
                {
                    newtb.Rows.Add(dtlotno.Rows[i].ItemArray);
                    lotnos = lotnos + "," + dtlotno.Rows[i]["LotNos"].ToString();
                }
            }

            sheet.Cells[4, 7] = lotnos.Trim(',').ToUpper();

            string ssql = "select approvalman, Auditingman from  IQC_Approval";
            DataSet ds = DbAccess.SelectBySql(ssql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                sheet.Cells[46, 9] = ds.Tables[0].Rows[0]["approvalman"].ToString();
                sheet.Cells[46, 13] = ds.Tables[0].Rows[0]["Auditingman"].ToString();
            }
            sheet.Cells[46, 20] = testuser;

            if (finaltotalstate.Contains("拒收"))
            {
                sheet.Cells[44, 12] = "√ 判退";
                sheet.Cells[44, 2] = "  合格";
            }
            else
            {
                sheet.Cells[44, 2] = "√ 合格";
                sheet.Cells[44, 12] = "  判退";
            }
        }

        private void gridView_DoubleClick(object sender, EventArgs e)
        {
          //string lotno = gridView.GetRowCellValue(gridView.FocusedRowHandle, "批次号").ToString();
          //string delivery = gridView.GetRowCellValue(gridView.FocusedRowHandle, "订单号").ToString();
          //string materialcode = gridView.GetRowCellValue(gridView.FocusedRowHandle, "物料编码").ToString();
          //string suppliers = gridView.GetRowCellValue(gridView.FocusedRowHandle, "客户").ToString();
          //string customer = gridView.GetRowCellValue(gridView.FocusedRowHandle, "供应商").ToString();
          //string time = gridView.GetRowCellValue(gridView.FocusedRowHandle, "测试日期").ToString();
          //  string suppliersorcustomer;
          //  if (suppliers == "")
          //      suppliersorcustomer = customer;
          //  else
          //      suppliersorcustomer = suppliers;                 
          // dt = SelectTestListRecord("退料测试汇总", "", "", "", "",delivery, "",materialcode, "", "", "", "", "").Tables[0];
          //  if (dt.Rows.Count > 0 && dt != null)
          //      Copy("", suppliersorcustomer,time,delivery, dt);
          //  else
          //      MessageBox.Show("无法导出Excel文件", "提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);
        }

        private string GetISO2859ReAc(string materialstate, int qty, string AQL)
        {

            string AR = "0/1";
            string ssql = "";
            DataTable dt = null;

            if (materialstate == "放宽检验")
            {
                ssql = " select cast(a.Ac as varchar(10))+'/'+cast(a.Re as varchar(10)) AR from IHPS_QUALITY_SPC_AQLIS02859 a left join IQC_TestSTD105ECheckSet i on i.Code1=a.Code left join IQC_TestSTD105ECode c on c.Code = a.Code where Type = '放宽'  and(LotSizemin <= " + qty + "  and LotSizemax >=" + qty + " and CheckLevel = 'II') and AQLValue =" + AQL + " ";
                dt = DbAccess.SelectBySql(ssql).Tables[0];
                if (dt != null && dt.Rows.Count > 0)
                {
                    AR = dt.Rows[0]["AR"].ToString();
                }
            }
            else if (materialstate == "正常检验")
            {
                ssql = " select cast(a.Ac as varchar(10))+'/'+cast(a.Re as varchar(10)) AR from IHPS_QUALITY_SPC_AQLIS02859 a left join IQC_TestSTD105ECheckSet i on i.Code=a.Code left join IQC_TestSTD105ECode c on c.Code = i.Code  where Type = '正常检验'  and(LotSizemin <=" + qty + " and LotSizemax >=" + qty + " and CheckLevel = 'II') and AQLValue = " + AQL + " ";
                DataSet ds = DbAccess.SelectBySql(ssql);
                dt = DbAccess.SelectBySql(ssql).Tables[0];
                if (dt != null && dt.Rows.Count > 0)
                {
                    AR = dt.Rows[0]["AR"].ToString();
                }
            }
            else if (materialstate == "加严检验")
            {
                ssql = " select cast(a.Ac as varchar(10))+'/'+cast(a.Re as varchar(10)) AR from IHPS_QUALITY_SPC_AQLIS02859 a left join IQC_TestSTD105ECheckSet i on i.Code=a.Code left join IQC_TestSTD105ECode c on c.Code = i.Code  where Type = '加严'  and(LotSizemin <=" + qty + " and LotSizemax >=" + qty + " and CheckLevel = 'II') and AQLValue = " + AQL + " ";
                DataSet ds = DbAccess.SelectBySql(ssql);
                dt = DbAccess.SelectBySql(ssql).Tables[0];
                if (dt != null && dt.Rows.Count > 0)
                {
                    AR = dt.Rows[0]["AR"].ToString();
                }
            }
            return AR;

        }

        public void Report(string sheetPrefixName, string sup, string dcode, string orderno,string lotno, DataTable tb,string returnqty,string returnreason,string Mdate,string ExpiryDate)
        {
            string finaltotalstate = "", testuser = "", AQLValue = "";
            int sheetCount = 1;
            if (sheetPrefixName == null || sheetPrefixName.Trim() == "")
                sheetPrefixName = " Sheet ";
            beforeTime = DateTime.Now;

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            afterTime = DateTime.Now;
            object missing = System.Reflection.Missing.Value;
            string templetFile = Environment.CurrentDirectory + @"\ReportFolder\ReturnReport.xls";
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(templetFile, missing, true, missing, missing, missing,
                                                          missing, missing, missing, missing, missing, missing, missing, missing, missing);
            Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets.get_Item(1);

            for (int i = 1; i < sheetCount; i++)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(i)).Copy(missing, workBook.Worksheets[i]);
            }

            #region  将源DataTable数据写入Excel

            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);
            sheet.Name = sheetPrefixName.Replace("/", "");

            string lotnos = "";
            DataTable dtlotno = null;

            sheet.Cells[4, 2] = sup;
            sheet.Cells[4, 18] = dcode;
            sheet.Cells[3, 17] = orderno;
            sheet.Cells[4, 12] = returnqty;
            sheet.Cells[4, 7] = lotno;
            sheet.Cells[3, 2] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            int qty =int.Parse(returnqty);
            sheet.Cells[7,2] = returnreason;
            sheet.Cells[7,12] = Mdate;
            sheet.Cells[7,19] = ExpiryDate;


            for (int j = 0; j < tb.Rows.Count; j++)
            {
               
                string finalstate = "";
                sheet.Cells[2, 1] = "IQC退料检验报表(" + tb.Rows[0 + j]["TestType"].ToString() + ")";
               // sheet.Cells[3, 2] = tb.Rows[0 + j]["TestTime"].ToString();
               // sheet.Cells[3, 6] = tb.Rows[0 + j]["lot_number"].ToString();
                               
                //sheet.Cells[4, 12] = tb.Rows[0 + j]["Qty"].ToString();
                //int qty = int.Parse(tb.Rows[0 + j]["Qty"].ToString());

                //sheet.Cells[3, 17] = tb.Rows[0 + j]["LotNo"].ToString();
                sheet.Cells[46, 19] = tb.Rows[0 + j]["TestUser"].ToString();
                sheet.Cells[5, 2] = tb.Rows[0 + j]["Productcode"].ToString();
                sheet.Cells[5, 7] = tb.Rows[0 + j]["materialenname"].ToString();
                sheet.Cells[10 + j + 1, 1] = tb.Rows[0 + j]["TestItem"].ToString();
                sheet.Cells[10 + j + 1, 2] = tb.Rows[0 + j]["TestSubItem"].ToString();
                sheet.Cells[10 + j + 1, 3] = tb.Rows[0 + j]["TestDes"].ToString();
                sheet.Cells[10 + j + 1, 9] = tb.Rows[0 + j]["TestTools"].ToString();
                sheet.Cells[10 + j + 1, 11] = tb.Rows[0 + j]["SampleType"].ToString();
                if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("MI"))
                {
                    sheet.Cells[10 + j + 1, 20] = "√";
                    AQLValue = "1.5";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("MA"))
                {
                    sheet.Cells[10 + j + 1, 19] = "√";
                    AQLValue = "0.65";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("CR=0.1"))
                {
                    sheet.Cells[10 + j + 1, 18] = "√";
                    AQLValue = "0.1";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.40"))
                {
                    sheet.Cells[10 + j + 1, 21] = "√";
                    AQLValue = "0.4";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.0"))
                {
                    sheet.Cells[10 + j + 1, 22] = "√";
                    AQLValue = "1";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("CR=0.01"))
                {
                    sheet.Cells[10 + j + 1, 23] = "√";
                    AQLValue = "0.01";
                }
                else if (tb.Rows[0 + j]["SampleType"].ToString().Contains("ISO2859-1"))
                {
                    if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.5"))
                    {
                        sheet.Cells[10 + j + 1, 20] = "√";
                        AQLValue = "1.5";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.65"))
                    {
                        sheet.Cells[10 + j + 1, 19] = "√";
                        AQLValue = "0.65";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.10"))
                    {
                        sheet.Cells[10 + j + 1, 18] = "√";
                        AQLValue = "0.1";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.40"))
                    {
                        sheet.Cells[10 + j + 1, 21] = "√";
                        AQLValue = "0.4";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.0"))
                    {
                        sheet.Cells[10 + j + 1, 22] = "√";
                        AQLValue = "1";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.010"))
                    {
                        sheet.Cells[10 + j + 1, 23] = "√";
                        AQLValue = "0.01";
                    }
                }
                else if (tb.Rows[0 + j]["SampleType"].ToString().Contains("C=0"))
                {
                    if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.5"))
                    {
                        sheet.Cells[10 + j + 1, 20] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.65"))
                    {
                        sheet.Cells[10 + j + 1, 19] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.1"))
                    {
                        sheet.Cells[10 + j + 1, 18] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.4"))
                    {
                        sheet.Cells[10 + j + 1, 21] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1"))
                    {
                        sheet.Cells[10 + j + 1, 22] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.01"))
                    {
                        sheet.Cells[10 + j + 1, 23] = "√";
                    }
                }

                if (tb.Rows[0 + j]["SampleType"].ToString().Contains("ISO2859-1"))
                {
                    sheet.Cells[10 + j + 1, 17] = GetISO2859ReAc("正常检验", qty, AQLValue);
                }
                else if (tb.Rows[0 + j]["SampleType"].ToString().Contains("C=0"))
                {
                    sheet.Cells[10 + j + 1, 17] = "0/1";
                }
                else
                {
                    sheet.Cells[10 + j + 1, 17] = GetReAc(tb.Rows[0 + j]["Sampleqty"].ToString(), tb.Rows[0 + j]["AQL"].ToString());
                }
 
                DataTable dlist = ic.SelectTestListRecord("退料测试明细", "", tb.Rows[j]["TestItem"].ToString(), tb.Rows[j]["TestSubItem"].ToString(), "", tb.Rows[j]["lotno"].ToString(), "", tb.Rows[j]["Productcode"].ToString(), "", lotno, "", "", "").Tables[0];

                dtlotno = dlist;
                string s = "";
                int testqty = 0;
                for (int m = 0; m < dlist.Rows.Count; m++)
                {
                    testqty = testqty + 1;
                    s = s + (m + 1).ToString() + ".测试值:" + dlist.Rows[0 + m]["Testvalue"].ToString() + ",检验结果:" + dlist.Rows[0 + m]["TestResult"].ToString() + ",备注:" + dlist.Rows[0 + m]["Remarks"].ToString() + ";";
                    finalstate = finalstate + dlist.Rows[0 + m]["TestFinalResult"].ToString();
                }
                finaltotalstate = finaltotalstate + finalstate;
                testuser = tb.Rows[j]["TestUser"].ToString();

                sheet.Cells[10 + j + 1, 12] = s.TrimEnd(';').TrimStart(';');
                if (finalstate.Contains("拒收"))

                    sheet.Cells[10 + j + 1, 24] = "NG";
                else

                    sheet.Cells[10 + j + 1, 24] = "OK";
            }
            DataTable newtb = dtlotno.Clone();
            newtb.Rows.Add(dtlotno.Rows[0].ItemArray);
            lotnos = lotnos + "," + dtlotno.Rows[0]["LotNos"].ToString();
            for (int i = 1; i < dtlotno.Rows.Count; i++)
            {
                bool flag = true;
                foreach (DataRow dr in newtb.Rows)
                {
                    if (dtlotno.Rows[i]["LotNos"].ToString() == dr["LotNos"].ToString())
                    {
                        flag = false;
                        continue;
                    }
                }
                if (flag)
                {
                    newtb.Rows.Add(dtlotno.Rows[i].ItemArray);
                    lotnos = lotnos + "," + dtlotno.Rows[i]["LotNos"].ToString();
                }
            }
            #endregion
         

            //string ssql = "select approvalman, Auditingman from  IQC_Approval";
            //DataSet ds = DbAccess.SelectBySql(ssql);
            //if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            //{
            //    sheet.Cells[46, 9] = ds.Tables[0].Rows[0]["approvalman"].ToString();
            //    sheet.Cells[46, 13] = ds.Tables[0].Rows[0]["Auditingman"].ToString();
            //}
            sheet.Cells[46, 20] = testuser;

            if (finaltotalstate.Contains("拒收"))
            {
                sheet.Cells[44, 12] = "√ 判退";
                sheet.Cells[44, 2] = "  合格";
            }
            else
            {
                sheet.Cells[44, 2] = "√ 合格";
                sheet.Cells[44, 12] = "  判退";
            }

        }

        public void returnReport(string sheetPrefixName, string sup, string dcode, string orderno,string lotno,string checkitem, DataTable tb, string returnqty, string returnreason, string Mdate, string ExpiryDate)
        {
            string finaltotalstate = "", testuser = "", AQLValue = "";
            int sheetCount = 1;
            if (sheetPrefixName == null || sheetPrefixName.Trim() == "")
                sheetPrefixName = "Sheet";
            beforeTime = DateTime.Now;

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            afterTime = DateTime.Now;
            object missing = System.Reflection.Missing.Value;
            string templetFile = Environment.CurrentDirectory + @"\ReportFolder\MaterialExpiryDate.xls";
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(templetFile, missing, true, missing, missing, missing,
                                                          missing, missing, missing, missing, missing, missing, missing, missing, missing);
            Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets.get_Item(1);

            for (int i = 1; i < sheetCount; i++)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(i)).Copy(missing, workBook.Worksheets[i]);
            }

            #region  将源DataTable数据写入Excel

            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets.get_Item(1);
            sheet.Name = sheetPrefixName.Replace("/", "");

            string lotnos = "";
            DataTable dtlotno = null;

            sheet.Cells[4, 2] = sup;
            sheet.Cells[4, 18] = dcode;
            sheet.Cells[6, 2] = lotno;
            sheet.Cells[3, 17] = orderno;
            sheet.Cells[4, 12] = returnqty;
            sheet.Cells[4, 7] = checkitem;
            sheet.Cells[3, 2] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            int qty = int.Parse(returnqty);
            sheet.Cells[7, 2] = "超期重检";
            //sheet.Cells[7, 12] = Mdate;
            sheet.Cells[7, 19] = ExpiryDate;


            for (int j = 0; j < tb.Rows.Count; j++)
            {
                string finalstate = "";
                sheet.Cells[2, 1] = "IQC超期重检报表(" + tb.Rows[0 + j]["TestType"].ToString() + ")";
                sheet.Cells[46, 19] = tb.Rows[0 + j]["TestUser"].ToString();
                sheet.Cells[5, 2] = tb.Rows[0 + j]["Productcode"].ToString();
                sheet.Cells[5, 7] = tb.Rows[0 + j]["materialenname"].ToString();
                sheet.Cells[10 + j + 1, 1] = tb.Rows[0 + j]["TestItem"].ToString();
                sheet.Cells[10 + j + 1, 2] = tb.Rows[0 + j]["TestSubItem"].ToString();
                sheet.Cells[10 + j + 1, 3] = tb.Rows[0 + j]["TestDes"].ToString();
                sheet.Cells[10 + j + 1, 9] = tb.Rows[0 + j]["TestTools"].ToString();
                sheet.Cells[10 + j + 1, 11] = tb.Rows[0 + j]["SampleType"].ToString();

                if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("MI"))
                {
                    sheet.Cells[10 + j + 1, 23] = "√";
                    AQLValue = "1.5";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("MA"))
                {
                    sheet.Cells[10 + j + 1, 21] = "√";
                    AQLValue = "0.65";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("CR=0.1"))
                {
                    sheet.Cells[10 + j + 1, 19] = "√";
                    AQLValue = "0.1";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.40"))
                {
                    sheet.Cells[10 + j + 1, 20] = "√";
                    AQLValue = "0.4";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.0"))
                {
                    sheet.Cells[10 + j + 1, 22] = "√";
                    AQLValue = "1";
                }
                else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("CR=0.01"))
                {
                    sheet.Cells[10 + j + 1, 18] = "√";
                    AQLValue = "0.01";
                }
                else if (tb.Rows[0 + j]["SampleType"].ToString().Contains("ISO2859-1"))
                {
                    if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.5"))
                    {
                        sheet.Cells[10 + j + 1, 23] = "√";
                        AQLValue = "1.5";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.65"))
                    {
                        sheet.Cells[10 + j + 1, 21] = "√";
                        AQLValue = "0.65";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.10"))
                    {
                        sheet.Cells[10 + j + 1, 19] = "√";
                        AQLValue = "0.1";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.40"))
                    {
                        sheet.Cells[10 + j + 1, 20] = "√";
                        AQLValue = "0.4";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.0"))
                    {
                        sheet.Cells[10 + j + 1, 22] = "√";
                        AQLValue = "1";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.010"))
                    {
                        sheet.Cells[10 + j + 1, 18] = "√";
                        AQLValue = "0.01";
                    }
                }
                else if (tb.Rows[0 + j]["SampleType"].ToString().Contains("C=0"))
                {
                    if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1.5"))
                    {
                        sheet.Cells[10 + j + 1, 23] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.65"))
                    {
                        sheet.Cells[10 + j + 1, 21] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.1"))
                    {
                        sheet.Cells[10 + j + 1, 19] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.4"))
                    {
                        sheet.Cells[10 + j + 1, 20] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=1"))
                    {
                        sheet.Cells[10 + j + 1, 22] = "√";
                    }
                    else if (tb.Rows[0 + j]["AQL"].ToString().StartsWith("AQL=0.01"))
                    {
                        sheet.Cells[10 + j + 1, 18] = "√";
                    }
                }

                if (tb.Rows[0 + j]["SampleType"].ToString().Contains("ISO2859-1"))
                {
                    sheet.Cells[10 + j + 1, 17] = GetISO2859ReAc("正常检验", qty, AQLValue);
                }
                else if (tb.Rows[0 + j]["SampleType"].ToString().Contains("C=0"))
                {
                    sheet.Cells[10 + j + 1, 17] = "0/1";
                }
                else
                {
                    sheet.Cells[10 + j + 1, 17] = GetReAc(tb.Rows[0 + j]["Sampleqty"].ToString(), tb.Rows[0 + j]["AQL"].ToString());
                }

                DataTable dlist = ic.SelectTestListRecord("超期重检明细", "", tb.Rows[j]["TestItem"].ToString(), tb.Rows[j]["TestSubItem"].ToString(), "", tb.Rows[j]["lotno"].ToString(), "", tb.Rows[j]["Productcode"].ToString(), "",checkitem, "", "", "").Tables[0];

                dtlotno = dlist;
                string s = "";
                int testqty = 0;
                for (int m = 0; m < dlist.Rows.Count; m++)
                {
                    testqty = testqty + 1;
                    s = s + (m + 1).ToString() + ".测试值:" + dlist.Rows[0 + m]["Testvalue"].ToString() + ",检验结果:" + dlist.Rows[0 + m]["TestResult"].ToString() + ",备注:" + dlist.Rows[0 + m]["Remarks"].ToString() + ";";
                    finalstate = finalstate + dlist.Rows[0 + m]["TestFinalResult"].ToString();
                }
                finaltotalstate = finaltotalstate + finalstate;
                testuser = tb.Rows[j]["TestUser"].ToString();

                sheet.Cells[10 + j + 1, 12] = s.TrimEnd(';').TrimStart(';');
                if (finalstate.Contains("拒收"))

                    sheet.Cells[10 + j + 1, 24] = "NG";
                else

                    sheet.Cells[10 + j + 1, 24] = "OK";
            }
            DataTable newtb = dtlotno.Clone();
            newtb.Rows.Add(dtlotno.Rows[0].ItemArray);
            lotnos = lotnos + "," + dtlotno.Rows[0]["LotNos"].ToString();
            for (int i = 1; i < dtlotno.Rows.Count; i++)
            {
                bool flag = true;
                foreach (DataRow dr in newtb.Rows)
                {
                    if (dtlotno.Rows[i]["LotNos"].ToString() == dr["LotNos"].ToString())
                    {
                        flag = false;
                        continue;
                    }
                }
                if (flag)
                {
                    newtb.Rows.Add(dtlotno.Rows[i].ItemArray);
                    lotnos = lotnos + "," + dtlotno.Rows[i]["LotNos"].ToString();
                }
            }
            #endregion
            sheet.Cells[46, 20] = testuser;

            if (finaltotalstate.Contains("拒收"))
            {
                sheet.Cells[44, 12] = "√ 判退";
                sheet.Cells[44, 2] = "  合格";
            }
            else
            {
                sheet.Cells[44, 2] = "√ 合格";
                sheet.Cells[44, 12] = "  判退";
            }

        }

        private void OpenFile(string filepath, string pdffile)
        {
            string filename = "";
            filename = pdffile;
            System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo();
            info.FileName = @filename;
            info.WorkingDirectory = filepath;
            info.Arguments = "";
            try
            {
                System.Diagnostics.Process.Start(info);
            }
            catch (System.ComponentModel.Win32Exception we)
            {
                MessageBox.Show(this, we.Message);
                return;
            }
        }

        private void gridView_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            DataTable dts = gridControl.DataSource as DataTable;
            if (dts == null || dts.Rows.Count < 1)
                return;
            
            if (gridView.RowCount < 1)
                return;

            if (e.Column.FieldName == "批次号")
            {
                Cursor.Current = Cursors.WaitCursor;
                string Mdate = "", ExpiryDate = "";
                string lotno = gridView.GetFocusedRowCellValue("批次号").ToString();
                string materialcode = gridView.GetFocusedRowCellValue("料号").ToString();
                string deliveryid = gridView.GetFocusedRowCellValue("接收单号").ToString();
                string checkitem = gridView.GetFocusedRowCellValue("抽检流水号").ToString();

                if (checkitem != "")
                    return;

                string sql = " select Mdate 生产日期,ExpiryDate 失效日期 from IQC_TestListReturn  where Mdate <> ''  and ExpiryDate <>''and Productcode = '" + materialcode + "'  and receptid = '" + deliveryid + "'  ";
                // DataSet ds = DbAccess.SelectBySql(sql);
                DataTable dtM = DbAccess.SelectBySql(sql).Tables[0];
                if (dtM != null && dtM.Rows.Count > 0)
                {
                    Mdate = dtM.Rows[0]["生产日期"].ToString();
                    ExpiryDate = dtM.Rows[0]["失效日期"].ToString();
                }
                dt = ic.SelectTestListRecord("退料测试汇总", "", "", "", "", deliveryid, "", materialcode, "", lotno, "", "", "").Tables[0];

                try
                {
                    Report(dt.Rows[0]["TestType"].ToString(), gridView.GetFocusedRowCellValue("供应商名称").ToString(), gridView.GetFocusedRowCellValue("测试日期").ToString(), deliveryid, lotno, dt, gridView.GetFocusedRowCellValue("重检数量").ToString(), gridView.GetFocusedRowCellValue("退料原因").ToString(), Mdate, ExpiryDate);
                }
                catch
                {
                    MessageBox.Show("生成完整报表失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Cursor.Current = Cursors.Default;
                    return;
                }
                Cursor.Current = Cursors.Default;
            }
            else if (e.Column.FieldName == "抽检流水号")
            {
                Cursor.Current = Cursors.WaitCursor;
                string Mdate = "", ExpiryDate = "";
                string lotno = gridView.GetFocusedRowCellValue("批次号").ToString();
                string materialcode = gridView.GetFocusedRowCellValue("料号").ToString();
                string deliveryid = gridView.GetFocusedRowCellValue("接收单号").ToString();
                string checkitem = gridView.GetFocusedRowCellValue("抽检流水号").ToString();
                string returnqty = gridView.GetFocusedRowCellValue("重检数量").ToString();
                if (checkitem == "")
                    return;

                string slotnolist = "";
                string[] values = lotno.Split(new char[1] { ',' });
                for (int i = 0; i < values.Length; i++)
                {
                    slotnolist += ",'" + values[i].ToString() + "'";
                }
                slotnolist = slotnolist.TrimStart(',');

                //string sql = " select Mdate 生产日期,ExpiryDate 失效日期 from IQC_TestListReturn  where Mdate <> ''  and ExpiryDate <>''and Productcode = '" + materialcode + "'  and receptid = '" + deliveryid + "'  ";
                string sql = "  select distinct convert(varchar(10),ExpiryDate,121) 有效日期  from materialRelation where reelid in (" + slotnolist + ")  ";
                DataTable dtM = DbAccess.SelectBySql(sql).Tables[0];
                if (dtM != null && dtM.Rows.Count > 0)
                {
                    for (int i = 0; i < dtM.Rows.Count; i++)
                    {
                       // Mdate += dtM.Rows[i]["生产日期"].ToString() + ";";
                        ExpiryDate += dtM.Rows[i]["有效日期"].ToString() + ";";
                    }
                    //Mdate = dtM.Rows[0]["生产日期"].ToString();
                    //ExpiryDate = dtM.Rows[0]["失效日期"].ToString();
                }
                dt = ic.SelectTestListRecord("超期重检汇总", "", "", "", "", deliveryid, "", materialcode, "", checkitem, returnqty, "", "").Tables[0];
                if (dt.Rows.Count < 1)
                    return;
                try
                {
                    returnReport(dt.Rows[0]["TestType"].ToString(), gridView.GetFocusedRowCellValue("供应商名称").ToString(), gridView.GetFocusedRowCellValue("测试日期").ToString(), deliveryid, lotno, checkitem, dt, gridView.GetFocusedRowCellValue("重检数量").ToString(), gridView.GetFocusedRowCellValue("退料原因").ToString(),"", ExpiryDate);
                }
                catch
                {
                    MessageBox.Show("生成完整报表失败", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Cursor.Current = Cursors.Default;
                    return;
                }
                Cursor.Current = Cursors.Default;
            }
            else if (e.Column.FieldName == "料号")
            {
                Cursor.Current = Cursors.WaitCursor;
                string materialcode = gridView.GetFocusedRowCellValue("料号").ToString();
                string deliveryid = gridView.GetFocusedRowCellValue("接收单号").ToString();
                string checkitem = gridView.GetFocusedRowCellValue("抽检流水号").ToString();
                if (checkitem == "")
                    return;

                string openfilepath = "\\\\QMSSVR\\TinReport"+"\\"+ materialcode+"\\"+checkitem;
                if (Directory.Exists(openfilepath))
                {
                    string[] pt = Directory.GetFiles(openfilepath);
                    pt = pt.Where(s => !s.EndsWith("Thumbs.db")).ToArray();
                    if (pt.Length == 0)
                        return;
                    string filename = Path.GetFileName(pt[0].ToString());
                    FileStream fs = new FileStream(openfilepath+"\\"+filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    BinaryReader reader = new BinaryReader(fs);
                    string fileclass = "";
                    try
                    {
                        for (int j = 0; j < 2; j++)
                        {
                            fileclass += reader.ReadByte().ToString();
                        }
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                    {
                        OpenFile(openfilepath, filename);
                        fs.Close();
                        fs.Dispose();
                        reader.Close();
                        Cursor.Current = Cursors.Default;
                        return;
                    }
                }  
            }

         }
        private void gridView_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            /*
            if (e.Column.FieldName == "检验结果")
            {
                GridCellInfo GridCell = e.Cell as GridCellInfo;
                if (GridCell.CellValue.ToString() == "NG")
                {                    
                    e.Appearance.BackColor = Color.Red;
                }
            }       
            */
        }

        private void gridView_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
            {
                return;
            }
            if (gridView.GetDataRow(e.RowHandle)["检验结果"].ToString() == "NG")
            {
                e.Appearance.BackColor = Color.Red;
            }
        }

        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "退料明细清单";
            int n = name.LastIndexOf(".") + 1;
            if (n > 0) name = name.Substring(n, name.Length - n);
            dlg.Title = "导出 " + title;
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
            if (DevExpress.XtraEditors.XtraMessageBox.Show("Do you want to open this file?", "Export To...", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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

        private void sBtnexport_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null)
                return;
            if (dt.Rows.Count <= 0) return;
            string fileName = ShowSaveFileDialog("Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx");
            if (fileName == string.Empty) return;
            ExportToEx(fileName, "xlsx", gridView);
            OpenFile(fileName);
        }

        private void cbExpiryDate_CheckedChanged(object sender, EventArgs e)
        {
            if (cbExpiryDate.Checked == true)
            {
                txtlotno.Text = "";
                txtlotno.Enabled = false;
            }
            else
            {
                txtlotno.Text = "";
                txtlotno.Enabled = true;
            }
        }

        private void gridView_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            //int hand = e.RowHandle;
            //if (hand < 0) return;
            //string testresult = gridView.GetRowCellValue(gridView.FocusedRowHandle, "检验结果").ToString();
            //if (testresult == null)
            //    return;                 
            //if (testresult == "NG")
            // {
            //        e.Appearance.BackColor = Color.Yellow;   // 改变行背景颜色
            // }
        }
    }
}