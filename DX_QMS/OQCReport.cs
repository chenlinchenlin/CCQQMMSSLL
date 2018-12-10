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
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.IO;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using System.IO;
using System.Diagnostics;
using DX_QMS.Common;

namespace DX_QMS
{
    public partial class OQCReport : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public OQCReport()
        {
            InitializeComponent();
        }
        private void bind(string types, ComboBoxEdit com)
        {
            string ssql = "select Definevalue from OQC_TypeDefine where Definetype='" + types + " ' order by sort ";
            DataTable dt = DbAccess.SelectBySql(ssql).Tables[0];
            com.Properties.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                com.Properties.Items.Add(row["Definevalue"]);
            }
        }
        private void OQCReport_Load(object sender, EventArgs e)
        {
            bind("客户",txtcustomer);
        }

        private void Btnselect_Click(object sender, EventArgs e)
        {
            string where = " where 1=1 " ,sql = "" ;
            string customer = txtcustomer.Text.Trim();
            string workno = txtworkno.Text.Trim();
            string boxsno = txtboxsno.Text.Trim();
            string begintime = begindate.Text;
            string endtime = enddate.Text;

            if (customer == "F1")
            {

                if (!string.IsNullOrEmpty(workno))
                {
                    where += " and workno = '" + workno + "' ";
                }
                if (!string.IsNullOrEmpty(boxsno))
                {
                    where += " and CartonNo = '" + boxsno + "' ";
                }

                if (!string.IsNullOrEmpty(begintime))
                {
                    where += " and checkdate >= '" + begintime + " 00:00:00 '";
                }
                if (!string.IsNullOrEmpty(endtime))
                {
                    where += " and checkdate <='" + endtime + " 23:59:59 '";
                }
                sql = @" select max(checkdate) 日期,SNnumber SN号,max(CartonNo) 箱号, max(VersionNO) 版本号 from OQC_SampleNewList ";
                sql += where + "  group by SNnumber order by SNnumber asc ";

                DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
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
            else
            {
                MessageBox.Show("没有该客户，该工单的抽样数据", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void Btnreset_Click(object sender, EventArgs e)
        {
            txtcustomer.Text = "";
            txtworkno.Text = "";
            shipmentqty.Text = "";
            txtboxsno.Text = "";
            txtsampleqty.Text = "";
            begindate.Text = "";
            enddate.Text = "";
            gridControl.DataSource = null;
        }

        string cutstr(string longstr)
        {
            string shortstr = "";
            string[] sArray = longstr.Split(new string[] { "-" }, StringSplitOptions.RemoveEmptyEntries);
            shortstr = sArray[1] + "-" + sArray[2];
            return shortstr;
        }

        void ReportF1(string customer ,string workno)
        {
            string sql = "", sqlstr = "";
            string where = " ";
            string begintime = begindate.Text;
            string endtime = enddate.Text;
            int sheetCount = 1;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            object missing = System.Reflection.Missing.Value;
            string templetFile = Environment.CurrentDirectory + @"\ReportFolder\OQCF1report.xlsx";
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

            sql = @" select item, checkdate, model,sendqty,sampleqty from OQC_TestListNew where customer='"+ customer + "' and workno='"+ workno + "' order by checkdate desc";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            if (dt == null || dt.Rows.Count < 0)
            {
                MessageBox.Show("不存在该工单号，无法生成出货报告","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }
            sheet.Cells.get_Range("O3").Value = dt.Rows[0]["checkdate"].ToString();
            sheet.Cells.get_Range("F62").Value = dt.Rows[0]["checkdate"].ToString();
            sheet.Cells.get_Range("P62").Value = dt.Rows[0]["checkdate"].ToString();
            sheet.Cells.get_Range("B5").Value = dt.Rows[0]["model"].ToString();

            sheet.Cells.get_Range("O4").Value = shipmentqty.Text.ToString();
            sheet.Cells.get_Range("O5").Value = txtsampleqty.Text.ToString();


            if (!string.IsNullOrEmpty(begintime))
            {
                where += " and checkdate >= '" + begintime + " 00:00:00 '";
            }
            if (!string.IsNullOrEmpty(endtime))
            {
                where += " and checkdate <='" + endtime + " 23:59:59 '";
            }

            sqlstr = @" select distinct CartonNo from ( select top "+ int.Parse( txtsampleqty.Text.ToString().Trim())+ " CartonNo from OQC_SampleNewList where workno = '"+ workno + "'"+ where+ " order by checkdate desc ) t  ";
            DataTable CartonNodt = DbAccess.SelectBySql(sqlstr).Tables[0];
            if (CartonNodt == null && CartonNodt.Rows.Count < 0)
            {
                MessageBox.Show("该工单号没有抽检，无法生成出货报告", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string CartonNoStr = "";
            CartonNoStr = CartonNodt.Rows[0]["CartonNo"].ToString();
            for (int i=1; i< CartonNodt.Rows.Count; i++)
            {
                CartonNoStr = CartonNoStr + "&" + cutstr(CartonNodt.Rows[i]["CartonNo"].ToString().Trim());
            }

            sheet.Cells.get_Range("B3").Value = CartonNoStr;


            sql = @" select top "+ int.Parse(txtsampleqty.Text.ToString().Trim())+ " * from OQC_SampleNewList where workno = '" + workno + "'"+ where+" order by checkdate desc  ";

           // sql = @" select * from OQC_SampleNewList where workno = '17110494-w'  order by checkdate desc  ";

            DataTable itemdt = DbAccess.SelectBySql(sql).Tables[0];
            if (itemdt == null && itemdt.Rows.Count < 0)
            {
                MessageBox.Show("该工单号没有抽检，无法生成出货报告", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            sheet.Cells.get_Range("B4").Value = "";
            sheet.Cells.get_Range("F5").Value = itemdt.Rows[0]["VersionNO"].ToString();

            int n = itemdt.Rows.Count / 50;
            int a = (int)'A';
            for (int i = 1, j = 1, m = 8; i <= itemdt.Rows.Count; i++)
            {
                //sheet.Cells[m, j] = i;
                sheet.Cells.get_Range((char)a + m.ToString()).Value = i;
                // sheet.Cells.get_Range((char)a+m.ToString()).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);    

                // sheet.Cells[m, j + 1] = itemdt.Rows[i-1]["SNnumber"].ToString().Trim();
                sheet.Cells.get_Range((char)(a + 1) + m.ToString()).Value = itemdt.Rows[i - 1]["SNnumber"].ToString().Trim();
                // sheet.Cells.get_Range((char)(a+1)+ m.ToString()).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);

                // sheet.Cells[m, j + 2] = itemdt.Rows[i-1]["CartonNo"].ToString().Trim();
                sheet.Cells.get_Range((char)(a + 2) + m.ToString()).Value = itemdt.Rows[i - 1]["CartonNo"].ToString().Trim();
                // sheet.Cells.get_Range((char)(a + 2)+ m.ToString()).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);

                //sheet.Cells[m, j + 3] = "ACC";
                sheet.Cells.get_Range((char)(a + 3) + m.ToString()).Value = "ACC";
                // sheet.Cells.get_Range((char)(a + 3)+ m.ToString()).Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Yellow);

                m++;
                if (i % 50 == 0 && i != 0)
                {
                    m = 8;
                    a = a + 4;
                    j = j + 4;
                }

            }

        }

        private void BtnReport_Click(object sender, EventArgs e)
        {
            if (txtcustomer.Text == "F1")
            {
                if (txtworkno.Text.Trim() == "")
                {
                    MessageBox.Show("请输入工单号", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (shipmentqty.Text == "")
                {
                    MessageBox.Show("请输入出货数量", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                int sendqty = 0;
                if (!int.TryParse(shipmentqty.Text, out sendqty))
                {
                    MessageBox.Show("送检数量不正确", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    shipmentqty.Text = "";
                    shipmentqty.Focus();
                    return;
                }
                if (int.Parse(shipmentqty.Text) < 0)
                {
                    MessageBox.Show("送检数量不能为负数", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    shipmentqty.Text = "";
                    shipmentqty.Focus();
                    return;
                }
                if (txtsampleqty.Text == "")
                {
                    MessageBox.Show("没有抽样数量", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                try
                {
                    ReportF1("F1", txtworkno.Text.Trim());
                }
                catch (Exception ex)
                {
                    MessageBox.Show("生成出货报告失败！","提醒",MessageBoxButtons.OK,MessageBoxIcon.Information);
                }


            }
            else if (txtcustomer.Text == "C1")
            {
                MessageBox.Show("暂时没有该客户的出货报告", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;

            }
            else
            {
                MessageBox.Show("暂时没有该客户的出货报告", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }


        }

        private void shipmentqty_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13 && shipmentqty.Text != "")
            {
                if (shipmentqty.Text == "")
                {
                    MessageBox.Show("请输入出货数量", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                int sendqty = 0;
                if (!int.TryParse(shipmentqty.Text, out sendqty))
                {
                    MessageBox.Show("送检数量不正确", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    shipmentqty.Text = "";
                    shipmentqty.Focus();
                    return;
                }
                if (int.Parse(shipmentqty.Text) < 0)
                {
                    MessageBox.Show("送检数量不能为负数", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    shipmentqty.Text = "";
                    shipmentqty.Focus();
                    return;
                }

                if (sendqty > 1)
                {
                    double AQLValue = 0.4;
                    if (sendqty >= 500000)
                    {
                        sendqty = 500000;
                    }
                    
                    string cosample = @"select case when sampleqty = '*' then '*' else sampleqty end as Sampleqty from IHPS_QUALITY_SPC_AQLC0 ";
                    cosample += " where ( Lowervalue <=" + sendqty + "and Uppervalue >=" + sendqty.ToString() + "and AQLValue = " + AQLValue + ")";
                
                    DataTable dtqty = DbAccess.SelectBySql(cosample).Tables[0];
                  
                    string SampleqtyMA = dtqty.Rows[0]["Sampleqty"].ToString();
                  
                    if (SampleqtyMA.Contains("*"))
                    {
                        txtsampleqty.Text = sendqty.ToString();
                    }
                    else
                    {
                        txtsampleqty.Text = SampleqtyMA;
                    }                    
                }
                else
                {
                    txtsampleqty.Text = "1";
                }
            }
        }
    }
}