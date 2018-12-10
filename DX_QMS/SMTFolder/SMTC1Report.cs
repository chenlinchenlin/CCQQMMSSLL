using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using DevExpress.XtraCharts;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DX_QMS.Common;
using DevExpress.Utils;
using System.Data.SqlClient;

namespace DX_QMS.SMTFolder
{
    public partial class SMTC1Report : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public SMTC1Report()
        {
            InitializeComponent();
        }


        public DataSet DailyCheck_CheckResultOper(string selecttype, string reporttype, string content, string startdate, string enddate)
        {
            SqlParameter[] para = new SqlParameter[5];
            para[0] = new SqlParameter("@selecttype", selecttype);
            para[1] = new SqlParameter("@reporttype", reporttype);
            para[2] = new SqlParameter("@content", content);
            para[3] = new SqlParameter("@startdate", startdate);
            para[4] = new SqlParameter("@enddate", enddate);
            return DbAccess.DataAdapterByCmd(CommandType.StoredProcedure, "QMS_SMT_Yield", para);
        }


        private void SMTC1Report_Load(object sender, EventArgs e)
        {
            txtreporttype.SelectedIndex = 0;
            txtstartdate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            txtenddate.Text = DateTime.Now.ToString("yyyy-MM-dd");
        }

        private void txtstartdate_EditValueChanged(object sender, EventArgs e)
        {
            if (selecttype.SelectedIndex == 2)
            {
                DataTable dt = null;
                string sql = "  select distinct customer from  OEM_MainTain  where lotno >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and lotno <=  '" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59' and lotno is not null and customer<>''  ";
                dt = DbAccess.SelectBySql(sql).Tables[0];
                checkcustomer.DataSource = dt;
                checkcustomer.DisplayMember = dt.Columns["customer"].ToString();
                checkcustomer.ValueMember = dt.Columns["customer"].ToString();

            }

        }

        private void selecttype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (selecttype.SelectedIndex == 0)
            {
                txtworkno.Enabled = true;
                lblworkno.Text = "多个工单逗号分割";
                lblmaterialcode.Text = "";
                txtmaterialcode.Text = "";
                txtmaterialcode.Enabled = false;
                checkcustomer.DataSource = null;
                checkcustomer.Enabled = false;

            }
            else if (selecttype.SelectedIndex == 1)
            {
                txtworkno.Text = "";
                txtworkno.Enabled = false;
                lblworkno.Text = "";
                lblmaterialcode.Text = "多个机型逗号分割";
                txtmaterialcode.Enabled = true;
                checkcustomer.DataSource = null;
                checkcustomer.Enabled = false;
            }
            else if (selecttype.SelectedIndex == 2)
            {
                txtworkno.Text = "";
                txtworkno.Enabled = false;
                txtmaterialcode.Text = "";
                txtmaterialcode.Enabled = false;
                checkcustomer.Enabled = true;
                DataTable dt = null;
                string sql = "   select distinct customer from  OEM_MainTain  where lotno >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and lotno <=  '" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59' and lotno is not null and customer<>''  ";
                dt = DbAccess.SelectBySql(sql).Tables[0];
                checkcustomer.DataSource = dt;
                checkcustomer.DisplayMember = dt.Columns["customer"].ToString();
                checkcustomer.ValueMember = dt.Columns["customer"].ToString();

            }

        }

        private void sBtnselect_Click(object sender, EventArgs e)
        {
            if (txtstartdate.Text == "" || txtenddate.Text == "")
            {
                MessageBox.Show("请输入相关的日期","提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);
                return;
            }
            if (selecttype.SelectedIndex == 0)
            {
                if (txtworkno.Text.Trim() == "")
                    return;
                DataTable dt = null;
                string workno = " 1=2 ";
                string str = txtworkno.Text.Trim();
                string[] sArray = str.Split(new char[4] { ';', '；', ',', '，' });
                for (int i = 0; i < sArray.Length; i++)
                {
                    if (sArray[i].Trim() == "")
                        continue;
                    workno += " or  workno ='" + sArray[i] + "'";
                }
                if (txtreporttype.Text == "日报")
                {
                    
                                        



                }
                else if (txtreporttype.Text == "周报")
                {
                   



                }
                else if (txtreporttype.Text == "月报")
                {
                    



                }

                gridControl.DataSource = null;
                gridView.Columns.Clear();
                gridControl.DataSource = dt;

                ChartTitle chartTitle = new ChartTitle();
                chartTitle.Text = txtstartdate.DateTime.ToString("yyyy-MM-dd") + "至" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " " + str + " 的良率情况";
                chartTitle.TextColor = System.Drawing.Color.Black; //颜色设置
                chartTitle.Font = new Font("微软雅黑", 10);
                chartTitle.Alignment = StringAlignment.Center;
                chartControl.Titles.Clear();
                chartControl.Titles.Add(chartTitle);
                chartControl.Series.Clear();

                for (int i = 0; i < sArray.Length; i++)
                {
                    if (sArray[i].Trim() == "")
                        continue;                   
                }
                ((XYDiagram)chartControl.Diagram).AxisY.Title.Visibility = DefaultBoolean.True;
                ((XYDiagram)chartControl.Diagram).AxisY.Title.Text = "良率";
                ((XYDiagram)chartControl.Diagram).AxisY.Color = Color.Gray;
                ((XYDiagram)chartControl.Diagram).AxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;
                AxisRange DIA = (AxisRange)((XYDiagram)chartControl.Diagram).AxisY.Range;
                DIA.SetMinMaxValues(0.9, 1);
            }
            else if (selecttype.SelectedIndex == 1)
            {




            }
            else if (selecttype.SelectedIndex == 2)
            {




            }

        }

        private void sBtnreport_Click(object sender, EventArgs e)
        {

        }

        private void sBtnreset_Click(object sender, EventArgs e)
        {




        }

        private void BtnBadClassDetails_Click(object sender, EventArgs e)
        {

        }

        private void sBtnExportBadClass_Click(object sender, EventArgs e)
        {

        }
    }
}