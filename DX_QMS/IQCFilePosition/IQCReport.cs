using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DX_QMS.Common;
using DevExpress.XtraCharts;
using DevExpress.Utils;
using DevExpress.XtraPrinting;
using DevExpress.XtraPrintingLinks;
using System.IO;
using DevExpress.XtraGrid.Views.Grid;

namespace DX_QMS.IQCFilePosition
{
    public partial class IQCReport : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        DataTable IQCtestreport = null;
        public IQCReport()
        {
            InitializeComponent();
        }


        public DataTable CreateTable()
        {
            DataTable namesTable = new DataTable("customer");
            DataColumn customer = new DataColumn();
            customer.DataType = System.Type.GetType("System.String");
            customer.ColumnName = "customer";
            namesTable.Columns.Add(customer);
         
            namesTable.Rows.Add("F1专用");
            namesTable.Rows.Add("X1专用");
            namesTable.Rows.Add("E1专用");
            namesTable.Rows.Add("T1专用");
            namesTable.Rows.Add("G2专用");
            namesTable.Rows.Add("H2专用");
            namesTable.Rows.Add("HBTG专用");
            namesTable.Rows.Add("Sokon专用");
            namesTable.Rows.Add("106004专用");
            namesTable.Rows.Add("B1专用");
            namesTable.Rows.Add("C1专用");

            return namesTable;

        }

        void bindEMScustomer()
        {
            DataTable dt = null;
            string sql = "  select customer from IQC_Customer where custometype = 'EMS客户' ";           
            dt = DbAccess.SelectBySql(sql).Tables[0];
            checkcustomer.DataSource = dt;
            checkcustomer.DisplayMember = dt.Columns["customer"].ToString();
            checkcustomer.ValueMember = dt.Columns["customer"].ToString();
        }

        void bindEMSCucustomer()
        {
            DataTable dt = null;
            string sql = "  select  CASE WHEN customer like '%客供%'  then  customer  ELSE  customer+'客供' end  客户 ";
                   sql += " from(select distinct customer from deliveryEMSOtherRec where customer <> ''  and(eventtime >= '"+txtstartdate.DateTime.ToString("yyyy-MM-dd HH:mm:ss")+ "' and eventtime <= '"+ txtenddate.DateTime.ToString("yyyy-MM-dd HH:mm:ss")+ "')) t   ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            checkcustomer.DataSource = dt;
            checkcustomer.DisplayMember = dt.Columns["客户"].ToString();
            checkcustomer.ValueMember = dt.Columns["客户"].ToString();
        }

        string selectEMScustomer()
        {
            string customer = "";
            for (int i = 0; i < checkcustomer.CheckedItems.Count; i++)
            {
                customer += "|"+checkcustomer.CheckedItems[i].ToString();
            }
            return customer;
        }


        private void IQCReport_Load(object sender, EventArgs e)
        {
            txtreporttype.SelectedIndex = 0;
            txtmaterialtype.SelectedIndex = 0;
            txtbusinessType.SelectedIndex = 0;

            txtstartdate.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            txtenddate.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }

        private void txtreporttype_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void txtbusinessType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtbusinessType.Text == "主营")
            {        
                checkcustomer.DataSource = null;
                checkcustomer.Enabled = false;
                txtmaterialtype.Enabled = true;
                txtmaterialtype.SelectedIndex = 0;
            }
            else if (txtbusinessType.Text == "EMS专用")
            {
                txtmaterialtype.Enabled = false;
                checkcustomer.Enabled = true;
                bindEMScustomer();
            }
            else if (txtbusinessType.Text == "EMS客供")
            {
                txtmaterialtype.Enabled = false;
                checkcustomer.Enabled = true;
                bindEMSCucustomer();
            }

        }

        private void txtstartdate_EditValueChanged(object sender, EventArgs e)
        {

           if (txtbusinessType.Text == "EMS客供")
            {
                txtmaterialtype.Enabled = false;
                checkcustomer.Enabled = true;
                bindEMSCucustomer();
            }
        }


        private void txtgoalrate_Leave(object sender, EventArgs e)
        {

            float output = 0;
            if (!float.TryParse(txtgoalrate.Text.Trim(), out output))
            {
                MessageBox.Show("目标达成率格式不正确", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtgoalrate.Text = "";
                txtgoalrate.Focus();
                return;
            }
            if (output <= 0 || output > 100)
            {
                MessageBox.Show("目标达成率在0到100之间", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtgoalrate.Text = "90";
                txtgoalrate.Focus();
                return;
            }

        }

        private string CheckString(string str)
        {
            string returnStr = "";
            if (str.IndexOf("'") != -1)
            {
                returnStr = str.Replace("'", "''");
                str = returnStr;
            }
            return str;
        }


        //private void sBtnselect_Click(object sender, EventArgs e)
        //{
        //    BackgroundTask.Backgroundselect(selectdata, null ,sBtnselect);
        //}


        private void sBtnselect_Click(object sender, EventArgs e)
        {
            string org = "";
            string reporttype = "";
            DataTable dt = null;
            if (txtstartdate.Text == "" || txtenddate.Text == "" )
            {
                MessageBox.Show("请选择起止时间和终止时间", "提醒",MessageBoxButtons.OK ,MessageBoxIcon.Information);
                return;
            }
            if (checkSHL.Checked == true && checkHCL.Checked == false)
            {
                org = " 组织 ='SHL' ";
            }
            else if (checkSHL.Checked == false && checkHCL.Checked == true)
            {
                org = " 组织 ='HCL' ";
            }
            else
            {
                org = " 组织 = 'SHL' or  组织 = 'HCL' ";
            }
            if (txtreporttype.Text == "周报")
            {
                reporttype = "";
            }

            if (txtreporttype.Text == "周报")
            {
                if (txtbusinessType.Text == "主营")
                {
                    gridControl.DataSource = null;
                    IQCreportchart.DataSource = null;

                    double rate = float.Parse(txtgoalrate.Text) * 0.01;

                    dt = weekzhu(org, txtmaterialtype.Text.Trim(),rate);

                    gridControl.DataSource = dt;
                    //(gridView.Columns[6].RealColumnEdit as RepositoryItemText).Mark.EditMark = "P3";
                    //(gridView.Columns[6].RealColumnEdit as RepositoryItemText).DisplayFormat.FormatString = "P3";
                    IQCreportchart.DataSource = dt;


                    int lotqty = 0, testqty = 0, NGqty = 0;
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            lotqty += int.Parse(dt.Rows[i]["来料批数"].ToString());
                            testqty += int.Parse(dt.Rows[i]["检验批数"].ToString());
                            NGqty += int.Parse(dt.Rows[i]["NG批数"].ToString());
                        }

                        txtcomelotqty.Text = lotqty.ToString();
                        txttestqty.Text = testqty.ToString();
                        txtNGqty.Text = NGqty.ToString();
                        txttestPassrate.Text = ((lotqty - NGqty + 0.0) / lotqty).ToString(); ;

                    }

                    ChartTitle chartTitle = new ChartTitle();
                    chartTitle.Text = txtstartdate.DateTime.ToString("yyyy-MM-dd") + "至" + txtenddate.DateTime.ToString("yyyy-MM-dd")+ txtmaterialtype.Text+"周报";
                    //chartTitle.TextColor = System.Drawing.Color.Black; //颜色设置
                    chartTitle.Font = new Font("Tahoma", 10);
                    chartTitle.Alignment = StringAlignment.Center;
                    IQCreportchart.Titles.Clear();
                    IQCreportchart.Titles.Add(chartTitle);

                    XYDiagram diagram = IQCreportchart.Diagram as XYDiagram;
                    IQCreportchart.SeriesTemplate.ArgumentDataMember = "时间";
                    diagram.AxisX.Title.Visibility = DefaultBoolean.True;
                    diagram.AxisX.Title.Text = "周数";
                    //diagram.AxisY.Title.Visibility = DefaultBoolean.True;
                    //diagram.AxisY.Title.Text = "数量";

                    IQCreportchart.Series["来料批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["来料批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["来料批数"].ValueDataMembers[0] = "来料批数";
                    IQCreportchart.Series["来料批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["来料批数"].LegendText = "来料批数";

                    IQCreportchart.Series["检验批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["检验批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["检验批数"].ValueDataMembers[0] = "检验批数";
                    IQCreportchart.Series["检验批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["检验批数"].LegendText = "检验批数";

                    IQCreportchart.Series["NG批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["NG批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["NG批数"].ValueDataMembers[0] = "NG批数";
                    IQCreportchart.Series["NG批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["NG批数"].LegendText = "NG批数";

                    ((XYDiagram)IQCreportchart.Diagram).SecondaryAxesY.Clear();
                    SecondaryAxisY myAxisY = new SecondaryAxisY();
                    ((XYDiagram)IQCreportchart.Diagram).SecondaryAxesY.Add(myAxisY);
                    myAxisY.Alignment = AxisAlignment.Far;
                    myAxisY.Title.Visibility = DefaultBoolean.True;
                    myAxisY.Title.Text = "百分率";
                    myAxisY.Color = Color.Gray;
                    myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;

                    IQCreportchart.Series["批合格率"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["批合格率"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["批合格率"].ValueDataMembers[0] = "批合格率";
                    IQCreportchart.Series["批合格率"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["批合格率"].LegendText = "批合格率";
                    ((LineSeriesView)IQCreportchart.Series["批合格率"].View).AxisY = myAxisY;
                    IQCreportchart.Series["批合格率"].PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;

                    IQCreportchart.Series["目标达成率"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["目标达成率"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["目标达成率"].ValueDataMembers[0] = "目标达成率";
                    IQCreportchart.Series["目标达成率"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["目标达成率"].LegendText = "目标达成率";
                    ((LineSeriesView)IQCreportchart.Series["目标达成率"].View).AxisY = myAxisY;
                    IQCreportchart.Series["目标达成率"].PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;

                }
                else if (txtbusinessType.Text == "EMS专用")
                {
                    gridControl.DataSource = null;
                    IQCreportchart.DataSource = null;

                    double rate = float.Parse(txtgoalrate.Text) * 0.01;

                    dt = weekEMS(org,rate);
                    gridControl.DataSource = dt;
                    IQCreportchart.DataSource = dt;


                    int lotqty = 0, testqty = 0, NGqty = 0;
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            lotqty += int.Parse(dt.Rows[i]["来料批数"].ToString());
                            testqty += int.Parse(dt.Rows[i]["检验批数"].ToString());
                            NGqty += int.Parse(dt.Rows[i]["NG批数"].ToString());
                        }

                        txtcomelotqty.Text = lotqty.ToString();
                        txttestqty.Text = testqty.ToString();
                        txtNGqty.Text = NGqty.ToString();                     
                        txttestPassrate.Text = ((lotqty - NGqty + 0.0) / lotqty).ToString(); ;

                    }

                    ChartTitle chartTitle = new ChartTitle();
                    chartTitle.Text = txtstartdate.DateTime.ToString("yyyy-MM-dd") + "至" + txtenddate.DateTime.ToString("yyyy-MM-dd")+ " EMS专用料周报";
                    //chartTitle.TextColor = System.Drawing.Color.Black; //颜色设置
                    chartTitle.Font = new Font("Tahoma", 10);
                    chartTitle.Alignment = StringAlignment.Center;
                    IQCreportchart.Titles.Clear();
                    IQCreportchart.Titles.Add(chartTitle);

                    XYDiagram diagram = IQCreportchart.Diagram as XYDiagram;
                    IQCreportchart.SeriesTemplate.ArgumentDataMember = "时间";
                    diagram.AxisX.Title.Visibility = DefaultBoolean.True;
                    diagram.AxisX.Title.Text = "周数";
                    //diagram.AxisY.Title.Visibility = DefaultBoolean.True;
                    //diagram.AxisY.Title.Text = "数量";

                    IQCreportchart.Series["来料批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["来料批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["来料批数"].ValueDataMembers[0] = "来料批数";
                    IQCreportchart.Series["来料批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["来料批数"].LegendText = "来料批数";

                    IQCreportchart.Series["检验批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["检验批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["检验批数"].ValueDataMembers[0] = "检验批数";
                    IQCreportchart.Series["检验批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["检验批数"].LegendText = "检验批数";

                    IQCreportchart.Series["NG批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["NG批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["NG批数"].ValueDataMembers[0] = "NG批数";
                    IQCreportchart.Series["NG批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["NG批数"].LegendText = "NG批数";

                    ((XYDiagram)IQCreportchart.Diagram).SecondaryAxesY.Clear();
                    SecondaryAxisY myAxisY = new SecondaryAxisY();
                    ((XYDiagram)IQCreportchart.Diagram).SecondaryAxesY.Add(myAxisY);
                    myAxisY.Alignment = AxisAlignment.Far;
                    myAxisY.Title.Visibility = DefaultBoolean.True;
                    myAxisY.Title.Text = "百分率";
                    myAxisY.Color = Color.Gray;
                    myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;

                    IQCreportchart.Series["批合格率"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["批合格率"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["批合格率"].ValueDataMembers[0] = "批合格率";
                    IQCreportchart.Series["批合格率"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["批合格率"].LegendText = "批合格率";
                    ((LineSeriesView)IQCreportchart.Series["批合格率"].View).AxisY = myAxisY;
                    IQCreportchart.Series["批合格率"].PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;

                    IQCreportchart.Series["目标达成率"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["目标达成率"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["目标达成率"].ValueDataMembers[0] = "目标达成率";
                    IQCreportchart.Series["目标达成率"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["目标达成率"].LegendText = "目标达成率";
                    ((LineSeriesView)IQCreportchart.Series["目标达成率"].View).AxisY = myAxisY;
                    IQCreportchart.Series["目标达成率"].PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;

                }
                else if (txtbusinessType.Text == "EMS客供")
                {

                    gridControl.DataSource = null;
                    IQCreportchart.DataSource = null;

                    double rate = float.Parse(txtgoalrate.Text) * 0.01;

                    dt = weekEMSCus(rate);
                    gridControl.DataSource = dt;
                    IQCreportchart.DataSource = dt;

                    int lotqty = 0, testqty = 0, NGqty = 0;
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            lotqty += int.Parse(dt.Rows[i]["来料批数"].ToString());
                            testqty += int.Parse(dt.Rows[i]["检验批数"].ToString());
                            NGqty += int.Parse(dt.Rows[i]["NG批数"].ToString());
                        }

                        txtcomelotqty.Text = lotqty.ToString();
                        txttestqty.Text = testqty.ToString();
                        txtNGqty.Text = NGqty.ToString();
                        txttestPassrate.Text = ((lotqty - NGqty + 0.0) / lotqty).ToString(); ;

                    }

                    ChartTitle chartTitle = new ChartTitle();
                    chartTitle.Text = txtstartdate.DateTime.ToString("yyyy-MM-dd") + "至" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " EMS客供料周报";
                    //chartTitle.TextColor = System.Drawing.Color.Black; //颜色设置
                    chartTitle.Font = new Font("Tahoma", 10);
                    chartTitle.Alignment = StringAlignment.Center;
                    IQCreportchart.Titles.Clear();
                    IQCreportchart.Titles.Add(chartTitle);

                    XYDiagram diagram = IQCreportchart.Diagram as XYDiagram;
                    IQCreportchart.SeriesTemplate.ArgumentDataMember = "时间";
                    diagram.AxisX.Title.Visibility = DefaultBoolean.True;
                    diagram.AxisX.Title.Text = "周数";
                    //diagram.AxisY.Title.Visibility = DefaultBoolean.True;
                    //diagram.AxisY.Title.Text = "数量";

                    IQCreportchart.Series["来料批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["来料批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["来料批数"].ValueDataMembers[0] = "来料批数";
                    IQCreportchart.Series["来料批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["来料批数"].LegendText = "来料批数";

                    IQCreportchart.Series["检验批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["检验批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["检验批数"].ValueDataMembers[0] = "检验批数";
                    IQCreportchart.Series["检验批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["检验批数"].LegendText = "检验批数";

                    IQCreportchart.Series["NG批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["NG批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["NG批数"].ValueDataMembers[0] = "NG批数";
                    IQCreportchart.Series["NG批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["NG批数"].LegendText = "NG批数";

                    ((XYDiagram)IQCreportchart.Diagram).SecondaryAxesY.Clear();
                    SecondaryAxisY myAxisY = new SecondaryAxisY();
                    ((XYDiagram)IQCreportchart.Diagram).SecondaryAxesY.Add(myAxisY);
                    myAxisY.Alignment = AxisAlignment.Far;
                    myAxisY.Title.Visibility = DefaultBoolean.True;
                    myAxisY.Title.Text = "百分率";
                    myAxisY.Color = Color.Gray;
                    myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;

                    IQCreportchart.Series["批合格率"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["批合格率"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["批合格率"].ValueDataMembers[0] = "批合格率";
                    IQCreportchart.Series["批合格率"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["批合格率"].LegendText = "批合格率";
                    ((LineSeriesView)IQCreportchart.Series["批合格率"].View).AxisY = myAxisY;
                    IQCreportchart.Series["批合格率"].PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;

                    IQCreportchart.Series["目标达成率"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["目标达成率"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["目标达成率"].ValueDataMembers[0] = "目标达成率";
                    IQCreportchart.Series["目标达成率"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["目标达成率"].LegendText = "目标达成率";
                    ((LineSeriesView)IQCreportchart.Series["目标达成率"].View).AxisY = myAxisY;
                    IQCreportchart.Series["目标达成率"].PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;

                }
            }
            else if (txtreporttype.Text == "月报")
            {
                if (txtbusinessType.Text == "主营")
                {
                    gridControl.DataSource = null;
                    IQCreportchart.DataSource = null;

                    double rate = float.Parse(txtgoalrate.Text) * 0.01;

                    dt = monthzhu(org,txtmaterialtype.Text.Trim(),rate);
                    gridControl.DataSource = dt;
                    IQCreportchart.DataSource = dt;


                    int lotqty = 0, testqty = 0, NGqty = 0;
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            lotqty += int.Parse(dt.Rows[i]["来料批数"].ToString());
                            testqty += int.Parse(dt.Rows[i]["检验批数"].ToString());
                            NGqty += int.Parse(dt.Rows[i]["NG批数"].ToString());
                        }
                        txtcomelotqty.Text = lotqty.ToString();
                        txttestqty.Text = testqty.ToString();
                        txtNGqty.Text = NGqty.ToString();
                        txttestPassrate.Text = ((lotqty - NGqty + 0.0) / lotqty).ToString(); ;
                    }

                    ChartTitle chartTitle = new ChartTitle();
                    chartTitle.Text = txtstartdate.DateTime.ToString("yyyy-MM-dd") + "至" + txtenddate.DateTime.ToString("yyyy-MM-dd") + txtmaterialtype.Text + "月报";
                    //chartTitle.TextColor = System.Drawing.Color.Black; //颜色设置
                    chartTitle.Font = new Font("Tahoma", 10);
                    chartTitle.Alignment = StringAlignment.Center;
                    IQCreportchart.Titles.Clear();
                    IQCreportchart.Titles.Add(chartTitle);



                    XYDiagram diagram = IQCreportchart.Diagram as XYDiagram;
                    IQCreportchart.SeriesTemplate.ArgumentDataMember = "时间";
                    diagram.AxisX.Title.Visibility = DefaultBoolean.True;
                    diagram.AxisX.Title.Text = "月份";


                    IQCreportchart.Series["来料批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["来料批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["来料批数"].ValueDataMembers[0] = "来料批数";
                    IQCreportchart.Series["来料批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["来料批数"].LegendText = "来料批数";

                    IQCreportchart.Series["检验批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["检验批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["检验批数"].ValueDataMembers[0] = "检验批数";
                    IQCreportchart.Series["检验批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["检验批数"].LegendText = "检验批数";

                    IQCreportchart.Series["NG批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["NG批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["NG批数"].ValueDataMembers[0] = "NG批数";
                    IQCreportchart.Series["NG批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["NG批数"].LegendText = "NG批数";


                    ((XYDiagram)IQCreportchart.Diagram).SecondaryAxesY.Clear();
                    SecondaryAxisY myAxisY = new SecondaryAxisY();
                    ((XYDiagram)IQCreportchart.Diagram).SecondaryAxesY.Add(myAxisY);
                    myAxisY.Alignment = AxisAlignment.Far;
                    myAxisY.Title.Visibility = DefaultBoolean.True;
                    myAxisY.Title.Text = "百分率";
                    myAxisY.Color = Color.Gray;
                    myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;


                    IQCreportchart.Series["批合格率"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["批合格率"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["批合格率"].ValueDataMembers[0] = "批合格率";
                    IQCreportchart.Series["批合格率"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["批合格率"].LegendText = "批合格率";
                    ((LineSeriesView)IQCreportchart.Series["批合格率"].View).AxisY = myAxisY;
                    IQCreportchart.Series["批合格率"].PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;

                    IQCreportchart.Series["目标达成率"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["目标达成率"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["目标达成率"].ValueDataMembers[0] = "目标达成率";
                    IQCreportchart.Series["目标达成率"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["目标达成率"].LegendText = "目标达成率";
                    ((LineSeriesView)IQCreportchart.Series["目标达成率"].View).AxisY = myAxisY;
                    IQCreportchart.Series["目标达成率"].PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;


                }
                else if (txtbusinessType.Text == "EMS专用")
                {                   
                    gridControl.DataSource = null;
                    IQCreportchart.DataSource = null;

                    double rate = float.Parse(txtgoalrate.Text) * 0.01;

                    dt = monthEMS(org ,rate);
                    gridControl.DataSource = dt;
                    IQCreportchart.DataSource = dt;



                    int lotqty = 0, testqty = 0, NGqty = 0;
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            lotqty += int.Parse(dt.Rows[i]["来料批数"].ToString());
                            testqty += int.Parse(dt.Rows[i]["检验批数"].ToString());
                            NGqty += int.Parse(dt.Rows[i]["NG批数"].ToString());
                        }
                        txtcomelotqty.Text = lotqty.ToString();
                        txttestqty.Text = testqty.ToString();
                        txtNGqty.Text = NGqty.ToString();
                        txttestPassrate.Text = ((lotqty - NGqty + 0.0) / lotqty).ToString(); ;
                    }

                    ChartTitle chartTitle = new ChartTitle();
                    chartTitle.Text = txtstartdate.DateTime.ToString("yyyy-MM-dd") + "至" + txtenddate.DateTime.ToString("yyyy-MM-dd")+ " EMS专用料月报";
                    //chartTitle.TextColor = System.Drawing.Color.Black; //颜色设置
                    chartTitle.Font = new Font("Tahoma", 10);
                    chartTitle.Alignment = StringAlignment.Center;
                    IQCreportchart.Titles.Clear();
                    IQCreportchart.Titles.Add(chartTitle);


                    XYDiagram diagram = IQCreportchart.Diagram as XYDiagram;
                    IQCreportchart.SeriesTemplate.ArgumentDataMember = "时间";
                    diagram.AxisX.Title.Visibility = DefaultBoolean.True;
                    diagram.AxisX.Title.Text = "月份";

                    IQCreportchart.Series["来料批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["来料批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["来料批数"].ValueDataMembers[0] = "来料批数";
                    IQCreportchart.Series["来料批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["来料批数"].LegendText = "来料批数";

                    IQCreportchart.Series["检验批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["检验批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["检验批数"].ValueDataMembers[0] = "检验批数";
                    IQCreportchart.Series["检验批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["检验批数"].LegendText = "检验批数";

                    IQCreportchart.Series["NG批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["NG批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["NG批数"].ValueDataMembers[0] = "NG批数";
                    IQCreportchart.Series["NG批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["NG批数"].LegendText = "NG批数";


                    ((XYDiagram)IQCreportchart.Diagram).SecondaryAxesY.Clear();
                    SecondaryAxisY myAxisY = new SecondaryAxisY();
                    ((XYDiagram)IQCreportchart.Diagram).SecondaryAxesY.Add(myAxisY);
                    myAxisY.Alignment = AxisAlignment.Far;
                    myAxisY.Title.Visibility = DefaultBoolean.True;
                    myAxisY.Title.Text = "百分率";
                    myAxisY.Color = Color.Gray;
                    myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;

                    IQCreportchart.Series["批合格率"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["批合格率"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["批合格率"].ValueDataMembers[0] = "批合格率";
                    IQCreportchart.Series["批合格率"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["批合格率"].LegendText = "批合格率";
                    ((LineSeriesView)IQCreportchart.Series["批合格率"].View).AxisY = myAxisY;
                    IQCreportchart.Series["批合格率"].PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;

                    IQCreportchart.Series["目标达成率"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["目标达成率"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["目标达成率"].ValueDataMembers[0] = "目标达成率";
                    IQCreportchart.Series["目标达成率"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["目标达成率"].LegendText = "目标达成率";
                    ((LineSeriesView)IQCreportchart.Series["目标达成率"].View).AxisY = myAxisY;
                    IQCreportchart.Series["目标达成率"].PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;

                }
                else if (txtbusinessType.Text == "EMS客供")
                {

                    gridControl.DataSource = null;
                    IQCreportchart.DataSource = null;
                   // IQCreportchart.ClearCache();
                   // IQCreportchart.Titles.Clear();
                   // IQCreportchart.Series.Clear();
                   //// IQCreportchart.DataFilters.Clear();
                   // IQCreportchart.RefreshData();
                   // IQCreportchart.DataBindings = null;              
                    double rate = float.Parse(txtgoalrate.Text) * 0.01;

                    dt = monthEMSCus(rate);
                    gridControl.DataSource = dt;

                    //IQCreportchart.DataBindings = dt;
                    IQCreportchart.DataSource = dt;                    

                    int lotqty = 0, testqty = 0, NGqty = 0;
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            lotqty += int.Parse(dt.Rows[i]["来料批数"].ToString());
                            testqty += int.Parse(dt.Rows[i]["检验批数"].ToString());
                            NGqty += int.Parse(dt.Rows[i]["NG批数"].ToString());
                        }
                        txtcomelotqty.Text = lotqty.ToString();
                        txttestqty.Text = testqty.ToString();
                        txtNGqty.Text = NGqty.ToString();
                        txttestPassrate.Text = ((lotqty - NGqty + 0.0) / lotqty).ToString(); ;
                    }

                    ChartTitle chartTitle = new ChartTitle();
                    chartTitle.Text = txtstartdate.DateTime.ToString("yyyy-MM-dd") + "至" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " EMS客供料月报";
                    //chartTitle.TextColor = System.Drawing.Color.Black; //颜色设置
                    chartTitle.Font = new Font("Tahoma", 10);
                    chartTitle.Alignment = StringAlignment.Center;
                    IQCreportchart.Titles.Clear();
                    IQCreportchart.Titles.Add(chartTitle);


                    XYDiagram diagram = IQCreportchart.Diagram as XYDiagram;
                    IQCreportchart.SeriesTemplate.ArgumentDataMember = "时间";
                    diagram.AxisX.Title.Visibility = DefaultBoolean.True;
                    diagram.AxisX.Title.Text = "月份";

                    IQCreportchart.Series["来料批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["来料批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["来料批数"].ValueDataMembers[0] = "来料批数";
                    IQCreportchart.Series["来料批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["来料批数"].LegendText = "来料批数";

                    IQCreportchart.Series["检验批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["检验批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["检验批数"].ValueDataMembers[0] = "检验批数";
                    IQCreportchart.Series["检验批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["检验批数"].LegendText = "检验批数";

                    IQCreportchart.Series["NG批数"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["NG批数"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["NG批数"].ValueDataMembers[0] = "NG批数";
                    IQCreportchart.Series["NG批数"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["NG批数"].LegendText = "NG批数";


                    ((XYDiagram)IQCreportchart.Diagram).SecondaryAxesY.Clear();
                    SecondaryAxisY myAxisY = new SecondaryAxisY();
                    ((XYDiagram)IQCreportchart.Diagram).SecondaryAxesY.Add(myAxisY);
                    myAxisY.Alignment = AxisAlignment.Far;
                    myAxisY.Title.Visibility = DefaultBoolean.True;
                    myAxisY.Title.Text = "百分率";
                    myAxisY.Color = Color.Gray;
                    myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;

                    IQCreportchart.Series["批合格率"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["批合格率"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["批合格率"].ValueDataMembers[0] = "批合格率";
                    IQCreportchart.Series["批合格率"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["批合格率"].LegendText = "批合格率";
                    ((LineSeriesView)IQCreportchart.Series["批合格率"].View).AxisY = myAxisY;
                    IQCreportchart.Series["批合格率"].PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;


                    IQCreportchart.Series["目标达成率"].ArgumentDataMember = "时间";
                    IQCreportchart.Series["目标达成率"].ArgumentScaleType = ScaleType.Qualitative;
                    IQCreportchart.Series["目标达成率"].ValueDataMembers[0] = "目标达成率";
                    IQCreportchart.Series["目标达成率"].LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
                    IQCreportchart.Series["目标达成率"].LegendText = "目标达成率";
                    ((LineSeriesView)IQCreportchart.Series["目标达成率"].View).AxisY = myAxisY;
                    IQCreportchart.Series["目标达成率"].PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;


                }
            }

        }

        DataTable weekzhu(string org_id ,string materialtype ,double rate)
        {
            string sqlcustomer = " 1=1 ", Oraclecustomer = "";
            string sqlmaterialtype = "";
            string oraclematerialtype = "";

            DataTable customerdt = null;
            string customersql = "  select customer from IQC_Customer where custometype = 'EMS客户' ";
            customerdt = DbAccess.SelectBySql(customersql).Tables[0];

            if (customerdt != null && customerdt.Rows.Count > 0)
            {
                for (int i = 0; i < customerdt.Rows.Count; i++)
                {
                    sqlcustomer += " and 名称 not like '%" + customerdt.Rows[i]["customer"].ToString() + "%'";
                    Oraclecustomer += "|" + customerdt.Rows[i]["customer"].ToString();
                }
                Oraclecustomer = Oraclecustomer.TrimStart('|');
            }


            if (materialtype == "主营电子料")
            {
                sqlmaterialtype += "  d.materialcode like '500[1-9]%' or d.materialcode like '501[1-5]%' or  d.materialcode like '510[134]%' or d.materialcode like '3201%' ";
                oraclematerialtype += " and (regexp_like(物料编码,'^500[1-9]')  or regexp_like(物料编码,'^501[1-5]') or regexp_like(物料编码,'^510[134]') or regexp_like(物料编码,'^3201[1-9]') ) ";
            }
            else if (materialtype == "主营结构料")
            {
                sqlmaterialtype += "  d.materialcode not like '500[1-9]%' or d.materialcode not like '501[1-5]%' or  d.materialcode not like '510[134]%' or d.materialcode not like '3201%' ";
                oraclematerialtype += " and ( not regexp_like(物料编码,'^500[1-9]')  or not regexp_like(物料编码,'^501[1-5]') or  not regexp_like(物料编码,'^510[134]') or not regexp_like(物料编码,'^3201[1-9]') ) ";
            }

            DataTable dt = null;
            string Oracle = " select to_char(处理日期,'yyyyIW') AS 周数,count(物料编码) as 来料批数 from (   ";
            Oracle += "  select * from APPS.CUX_INVIQC_DAY_V where not regexp_like(物料描述,'"+Oraclecustomer+ "') and  ";
            Oracle += "  (处理日期  between to_date('"+ txtstartdate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' , 'yyyy-mm-dd hh24:mi:ss') and to_date('"+txtenddate.DateTime.ToString("yyyy-MM-dd HH:mm:ss")+ "' , 'yyyy-mm-dd hh24:mi:ss') ) "+ oraclematerialtype + "  ) t   ";
            Oracle += "   WHERE ("+org_id+ ") and 处理类型 ='接收'  group by to_char(处理日期, 'yyyyIW') ORDER BY 周数   ";
            Oracle = CheckString(Oracle);
            string sql = "   select dbo.getWeekNoBtDate(测试日期) as 周数 ,count(*) 检验批数 ,COUNT (case when 检验结果='NG' then 'NG'  else null end ) as NG批数 from ";
                   sql += "  (select d.deliveryid 接收单号,d.materialcode 料号,max(materialname) 名称,case when max(testqty)>1 then max(testqty) else sum(cast(qty as bigint)) end 数量, ";
                   sql += "  case max(TestFinalResult) when '拒收' then 'NG' else 'OK' end 检验结果,max(testtime) 测试日期,MAX (case when org_id='HCL' then 'HCL' else 'SHL' end) 组织 ";
                   sql += "  from delivery d left join MaterialSpec m on d.materialcode=m.materialcode ";
                   sql += "  right join (select receptid,Productcode,LotNo,max(testqty) testqty,max(TestFinalResult) TestFinalResult,min(testtime) testtime from IQC_TestList ";
                   sql += "  where testtime >= '"+ txtstartdate.DateTime.ToString("yyyy-MM-dd HH:mm:ss")+ "'  and testtime <= '"+txtenddate.DateTime.ToString("yyyy-MM-dd HH:mm:ss")+ "' and receptid is not null ";
                   sql += "  group by receptid,Productcode,Lotno) x on d.deliveryid=x.receptid and d.materialcode=x.Productcode and d.lotno=x.LotNo ";
                   sql += "  where "+ sqlmaterialtype ;
                   sql += "  group by d.deliveryid,d.materialcode ) z where ("+org_id+ ") and ( "+sqlcustomer+ " ) ";
                   sql += "  group by dbo.getWeekNoBtDate(测试日期) ";
     string Oraclesql = " select t.周数 时间,l.来料批数,t.检验批数 ,t.NG批数 ,";
                   Oraclesql += " convert(numeric(5,3),(l.来料批数-t.NG批数)/(l.来料批数) ) as 批合格率,"+rate+ " as 目标达成率  from  (";
                   Oraclesql += sql+ " ) t  inner join ( select 周数,来料批数 FROM OPENQUERY(LINKERP,'"+ Oracle+ " ')) l on t.周数 = l.周数  order by  cast(t.周数 as int) asc  ";
            dt = DbAccess.SelectBySql(Oraclesql).Tables[0];
            return dt;


        }


        DataTable monthzhu(string org_id,string materialtype,double rate)
        {
            string sqlcustomer = " 1=1 ", Oraclecustomer = "";
            string sqlmaterialtype = "";
            string oraclematerialtype = "";

            DataTable customerdt = null;
            string customersql = "  select customer from IQC_Customer where custometype = 'EMS客户' ";
            customerdt = DbAccess.SelectBySql(customersql).Tables[0];

            if (customerdt != null && customerdt.Rows.Count > 0)
            {
                for (int i = 0; i < customerdt.Rows.Count; i++)
                {
                    sqlcustomer += " and 名称 not like '%" + customerdt.Rows[i]["customer"].ToString() + "%'";
                    Oraclecustomer += "|" + customerdt.Rows[i]["customer"].ToString();
                }
                Oraclecustomer = Oraclecustomer.TrimStart('|');
            }


            if (materialtype == "主营电子料")
            {
                sqlmaterialtype += "  d.materialcode like '500[1-9]%' or d.materialcode like '501[1-5]%' or  d.materialcode like '510[134]%' or d.materialcode like '3201%' ";
                oraclematerialtype += " and (regexp_like(物料编码,'^500[1-9]')  or regexp_like(物料编码,'^501[1-5]') or regexp_like(物料编码,'^510[134]') or regexp_like(物料编码,'^3201[1-9]') ) ";
            }
            else if (materialtype == "主营结构料")
            {
                sqlmaterialtype += "  d.materialcode not like '500[1-9]%' or d.materialcode not like '501[1-5]%' or  d.materialcode not like '510[134]%' or d.materialcode not like '3201%' ";
                oraclematerialtype += " and ( not regexp_like(物料编码,'^500[1-9]')  or not regexp_like(物料编码,'^501[1-5]') or  not regexp_like(物料编码,'^510[134]') or not regexp_like(物料编码,'^3201[1-9]') ) ";
            }

            DataTable dt = null;
            string Oracle = " select to_char(处理日期,'yyyyMM') AS 月份,count(物料编码) as 来料批数 from (   ";
            Oracle += "  select * from APPS.CUX_INVIQC_DAY_V where not regexp_like(物料描述,'"+Oraclecustomer+ "') and  ";
            Oracle += "  (处理日期  between to_date('" + txtstartdate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' , 'yyyy-mm-dd hh24:mi:ss') and to_date('" + txtenddate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' , 'yyyy-mm-dd hh24:mi:ss') ) " + oraclematerialtype + "  ) t   ";
            Oracle += "   WHERE (" + org_id + ") and 处理类型 ='接收'  group by to_char(处理日期, 'yyyyMM') ORDER BY 月份   ";
            Oracle = CheckString(Oracle);
            string sql = "  select dbo.getMonthByDate(测试日期) as 月份 ,count(*) 检验批数 ,COUNT (case when 检验结果='NG' then 'NG'  else null end ) as NG批数 from ";
            sql += "  (select d.deliveryid 接收单号,d.materialcode 料号,max(materialname) 名称,case when max(testqty)>1 then max(testqty) else sum(cast(qty as bigint)) end 数量, ";
            sql += "  case max(TestFinalResult) when '拒收' then 'NG' else 'OK' end 检验结果,max(testtime) 测试日期,MAX (case when org_id='HCL' then 'HCL' else 'SHL' end) 组织 ";
            sql += "  from delivery d left join MaterialSpec m on d.materialcode=m.materialcode ";
            sql += "  right join (select receptid,Productcode,LotNo,max(testqty) testqty,max(TestFinalResult) TestFinalResult,min(testtime) testtime from IQC_TestList ";
            sql += "  where testtime >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "'  and testtime <= '" + txtenddate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and receptid is not null ";
            sql += "  group by receptid,Productcode,Lotno) x on d.deliveryid=x.receptid and d.materialcode=x.Productcode and d.lotno=x.LotNo ";
            sql += "  where " + sqlmaterialtype;
            sql += "  group by d.deliveryid,d.materialcode ) z where (" + org_id + ") and ( "+sqlcustomer+ " ) ";
            sql += "  group by  dbo.getMonthByDate(测试日期) ";
            string Oraclesql = " select t.月份 时间,l.来料批数,t.检验批数 ,t.NG批数 ,";
            Oraclesql += " convert(numeric(5,3),(l.来料批数-t.NG批数)/(l.来料批数) ) as 批合格率,"+rate+" as 目标达成率  from  (";
            Oraclesql += sql + " ) t  inner join ( select 月份,来料批数 FROM OPENQUERY(LINKERP,'" + Oracle + " ')) l on t.月份 = l.月份 order by  cast (t.月份 as int) asc ";
            dt = DbAccess.SelectBySql(Oraclesql).Tables[0];
            return dt;

        }


        DataTable weekEMS(string org_id ,double rate)
        {
            string sqlcustomer = " 1=2" , Oraclecustomer = "";

            for (int i = 0; i < checkcustomer.CheckedItems.Count; i++)
            {
                sqlcustomer += " or materialname like '%" + checkcustomer.CheckedItems[i].ToString() + "%'"; 
                Oraclecustomer += "|" + checkcustomer.CheckedItems[i].ToString();
            }
            Oraclecustomer = Oraclecustomer.TrimStart('|');
            

            if (checkcustomer.CheckedItems.Count < 1 )
            {

                DataTable customerdt = null;
                string customersql = "  select customer from IQC_Customer where custometype = 'EMS客户' ";
                customerdt = DbAccess.SelectBySql(customersql).Tables[0];

                if (customerdt != null && customerdt.Rows.Count > 0)
                {
                    for (int i = 0; i < customerdt.Rows.Count; i++)
                    {
                        sqlcustomer += " or materialname like '%" + customerdt.Rows[i]["customer"].ToString() + "%'";
                        Oraclecustomer += "|" + customerdt.Rows[i]["customer"].ToString();
                    }
                    Oraclecustomer = Oraclecustomer.TrimStart('|');
                }
                //sqlcustomer = "   materialname like '%E1专用%' or materialname like '%F1专用%'  or materialname like '%T1专用%' or materialname like '%HBTG专用%' or materialname like '%G2专用%'  ";
                //sqlcustomer += "  or materialname like '%X1专用%' or materialname like '%H2专用%' or materialname like '%106004专用%'  or materialname like '%Sokon专用%' or materialname like '%B1专用%' or materialname like '%C1专用%' ";
                //Oraclecustomer += "E1专用|F1专用|HBTG专用|G2专用|106004专用|T1专用|X1专用|H2专用|Sokon专用|B1专用|C1专用";
            }
          
            DataTable dt = null;
            string Oracle = " select to_char(处理日期,'yyyyIW') AS 周数,count(物料编码) as 来料批数 from (   ";
            Oracle += "  select * from APPS.CUX_INVIQC_DAY_V where  regexp_like(物料描述,'"+Oraclecustomer+ "') and  ";
            Oracle += "  (处理日期  between to_date('" + txtstartdate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' , 'yyyy-mm-dd hh24:mi:ss') and to_date('" + txtenddate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' , 'yyyy-mm-dd hh24:mi:ss') ) ) t   ";
            Oracle += "   WHERE (" + org_id + ") and 处理类型 ='接收'  group by to_char(处理日期, 'yyyyIW') ORDER BY 周数   ";
            Oracle = CheckString(Oracle);
            string sql = "   select dbo.getWeekNoBtDate(测试日期) as 周数 ,count(*) 检验批数 ,COUNT (case when 检验结果='NG' then 'NG'  else null end ) as NG批数 from ";
            sql += "  (select d.deliveryid 接收单号,d.materialcode 料号,max(materialname) 名称,case when max(testqty)>1 then max(testqty) else sum(cast(qty as bigint)) end 数量,";
            sql += "  case max(TestFinalResult) when '拒收' then 'NG' else 'OK' end 检验结果,max(testtime) 测试日期,MAX (case when org_id='HCL' then 'HCL' else 'SHL' end) 组织  ";
            sql += "  from delivery d left join MaterialSpec m on d.materialcode=m.materialcode ";
            sql += "  right join (select receptid,Productcode,LotNo,max(testqty) testqty,max(TestFinalResult) TestFinalResult,min(testtime) testtime from IQC_TestList ";
            sql += "  where testtime >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "'  and testtime <= '" + txtenddate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and receptid is not null ";
            sql += "  group by receptid,Productcode,Lotno) x on d.deliveryid=x.receptid and d.materialcode=x.Productcode and d.lotno=x.LotNo ";
            sql += "  where ("+ sqlcustomer+")";
            sql += "  group by d.deliveryid,d.materialcode ) z where (" + org_id + ") group by dbo.getWeekNoBtDate(测试日期)  ";

            string Oraclesql = " select t.周数 时间,l.来料批数,t.检验批数 ,t.NG批数 ,";
            Oraclesql += " convert(numeric(5,3),(l.来料批数-t.NG批数)/(l.来料批数) ) as 批合格率,"+rate+ " as 目标达成率  from  (";
            Oraclesql += sql + " ) t  inner join ( select 周数,来料批数 FROM OPENQUERY(LINKERP,'" + Oracle + " ')) l on t.周数 = l.周数  order by  cast(t.周数 as int) asc  ";
            dt = DbAccess.SelectBySql(Oraclesql).Tables[0];
            return dt;
        }


        DataTable monthEMS(string org_id,double rate )
        {

            string sqlcustomer = "1=2", Oraclecustomer = "";

            for (int i = 0; i < checkcustomer.CheckedItems.Count; i++)
            {
                sqlcustomer += " or materialname like '%" + checkcustomer.CheckedItems[i].ToString() + "%'";
                Oraclecustomer += "|" + checkcustomer.CheckedItems[i].ToString();
            }
            Oraclecustomer = Oraclecustomer.TrimStart('|');

            if (checkcustomer.CheckedItems.Count < 1)
            {

                DataTable customerdt = null;
                string customersql = "  select customer from IQC_Customer where custometype = 'EMS客户' ";
                customerdt = DbAccess.SelectBySql(customersql).Tables[0];

                if (customerdt != null && customerdt.Rows.Count > 0)
                {
                    for (int i = 0; i < customerdt.Rows.Count; i++)
                    {
                        sqlcustomer += " or materialname like '%" + customerdt.Rows[i]["customer"].ToString() + "%'";
                        Oraclecustomer += "|" + customerdt.Rows[i]["customer"].ToString();
                    }
                    Oraclecustomer = Oraclecustomer.TrimStart('|');
                }

                //sqlcustomer = "   materialname like '%E1专用%' or materialname like '%F1专用%'  or materialname like '%T1专用%' or materialname like '%HBTG专用%' or materialname like '%G2专用%'  ";
                //sqlcustomer += "  or materialname like '%X1专用%' or materialname like '%H2专用%' or materialname like '%106004专用%' or materialname like '%Sokon专用%' or materialname like '%B1专用%' or materialname like '%C1专用%'  ";
                //Oraclecustomer += "E1专用|F1专用|HBTG专用|G2专用|106004专用|T1专用|X1专用|H2专用|Sokon专用|B1专用|C1专用";

            }

            DataTable dt = null;
            string Oracle = " select to_char(处理日期,'yyyyMM') AS 月份,count(物料编码) as 来料批数 from (   ";
            Oracle += "  select * from APPS.CUX_INVIQC_DAY_V where  regexp_like(物料描述,'" + Oraclecustomer + "') and  ";
            Oracle += "  (处理日期  between to_date('" + txtstartdate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' , 'yyyy-mm-dd hh24:mi:ss') and to_date('" + txtenddate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' , 'yyyy-mm-dd hh24:mi:ss') ) ) t   ";
            Oracle += "   WHERE (" + org_id + ") and 处理类型 ='接收'  group by to_char(处理日期, 'yyyyMM') ORDER BY 月份   ";
            Oracle = CheckString(Oracle);
            string sql = "  select dbo.getMonthByDate(测试日期) as 月份 ,count(*) 检验批数,COUNT (case when 检验结果='NG' then 'NG'  else null end ) as NG批数 from ";
            sql += "  (select d.deliveryid 接收单号,d.materialcode 料号,max(materialname) 名称,case when max(testqty)>1 then max(testqty) else sum(cast(qty as bigint)) end 数量,";
            sql += "  case max(TestFinalResult) when '拒收' then 'NG' else 'OK' end 检验结果,max(testtime) 测试日期,MAX (case when org_id='HCL' then 'HCL' else 'SHL' end) 组织  ";
            sql += "  from delivery d left join MaterialSpec m on d.materialcode=m.materialcode ";
            sql += "  right join (select receptid,Productcode,LotNo,max(testqty) testqty,max(TestFinalResult) TestFinalResult,min(testtime) testtime from IQC_TestList ";
            sql += "  where testtime >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "'  and testtime <= '" + txtenddate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and receptid is not null ";
            sql += "  group by receptid,Productcode,Lotno) x on d.deliveryid=x.receptid and d.materialcode=x.Productcode and d.lotno=x.LotNo ";
            sql += "  where (" + sqlcustomer + ")";
            sql += "  group by d.deliveryid,d.materialcode ) z where (" + org_id + ") group by dbo.getMonthByDate(测试日期)  ";

            string Oraclesql = " select t.月份 时间,l.来料批数,t.检验批数 ,t.NG批数,";
            Oraclesql += " convert(numeric(5,3),(l.来料批数-t.NG批数)/(l.来料批数) ) as 批合格率,"+rate+ " as 目标达成率  from  (";
            Oraclesql += sql + " ) t  inner join ( select 月份,来料批数 FROM OPENQUERY(LINKERP,'" + Oracle + " ')) l on t.月份 = l.月份  order by  cast(t.月份 as int) asc  ";
            dt = DbAccess.SelectBySql(Oraclesql).Tables[0];
            return dt;
        }


        DataTable weekEMSCus(double rate)
        {

            string sqlcustomer = "1=2";

            for (int i = 0; i < checkcustomer.CheckedItems.Count; i++)
            {
                sqlcustomer += " or r.customer like '%" + checkcustomer.CheckedItems[i].ToString().Replace("客供", "") + "%'";             
            }
            if (checkcustomer.CheckedItems.Count < 1)
            {
                sqlcustomer = "1=1 ";
            }

            DataTable dt = null;
            string sql = "  select t.周数 时间,l.来料批数,t.检验批数,t.NG批数, convert(numeric(5,3),(l.来料批数+0.0 - t.NG批数)/(l.来料批数) ) as 批合格率,"+rate+ " as 目标达成率 ";
            sql += "  from (  select dbo.getWeekNoBtDate(测试日期) as 周数 ,count(*) 检验批数 ,COUNT (case when 检验结果='NG' then 'NG'  else null end ) as NG批数 from   ";
            sql += "  (select d.deliveryid 接收单号,d.materialcode 料号,case max(TestFinalResult) when '拒收' then 'NG' else 'OK' end 检验结果,max(testtime) 测试日期  ";
            sql += "  from delivery d left join deliveryEMSOtherRec r on d.deliveryid=r.deliveryid    ";
            sql += "  right join (select receptid,Productcode,LotNo,max(testqty) testqty,max(TestFinalResult) TestFinalResult,min(testtime) testtime from IQC_TestList   ";
            sql += "  where  testtime >= '"+ txtstartdate.DateTime.ToString("yyyy-MM-dd HH:mm:ss")+ "'  and testtime <= '"+txtenddate.DateTime.ToString("yyyy-MM-dd HH:mm:ss")+ "' and receptid is not null  ";
            sql += "  group by receptid,Productcode,Lotno) x on d.deliveryid=x.receptid and d.materialcode=x.Productcode and d.lotno=x.LotNo   ";
            sql += "  where  d.materialcode like '7%' and ("+sqlcustomer+") group by d.deliveryid,d.materialcode  ";
            sql += "  ) z group by dbo.getWeekNoBtDate(测试日期) ) t  inner join  (  ";
            sql += "  select dbo.getWeekNoBtDate(来料日期) as 周数 ,count(HYT编码) 来料批数  from  ( ";
            sql += "  select  r.deliveryid 单号,m.materialcode  HYT编码, m.materialname 物料描述,r.vendorname 客户 ,r.arrivatedate 来料日期 ";
            sql += "  from deliveryEMSOtherRec r left join OEM_EMSHYTCusRelation cr on r.vendorcode=cr.cuscode and r.customer=cr.customer  ";
            sql += "  left join (select deliveryid,materialcode,min(transactiondate) eventtime  from delivery group by deliveryid,materialcode) d on cr.hytcode=d.materialcode and r.deliveryid=d.deliveryid ";
            sql += "  left join MaterialSpec m on cr.hytcode=m.materialcode  ";
            sql += "  where r.arrivatedate>='"+ txtstartdate.DateTime.ToString("yyyy-MM-dd HH:mm:ss")+ "' and r.arrivatedate<= '"+ txtenddate.DateTime.ToString("yyyy-MM-dd HH:mm:ss")+ "' and ("+sqlcustomer+ ") ";
            sql += "   ) la  group by dbo.getWeekNoBtDate(来料日期) ) l on t.周数 = l.周数  order by  cast (t.周数 as int) asc  ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;

        }

        DataTable monthEMSCus(double rate)
        {

            string sqlcustomer = "1=2";

            for (int i = 0; i < checkcustomer.CheckedItems.Count; i++)
            {
                sqlcustomer += " or r.customer like '%" + checkcustomer.CheckedItems[i].ToString().Replace("客供", "") + "%'";
            }
            if (checkcustomer.CheckedItems.Count < 1)
            {
                sqlcustomer = "1=1 ";
            }

            DataTable dt = null;
            string sql = "  select t.月份 时间,l.来料批数,t.检验批数,t.NG批数, convert(numeric(5,3),(l.来料批数+0.0 - t.NG批数)/(l.来料批数) ) as 批合格率,"+rate+ " as 目标达成率 ";
            sql += "  from (  select dbo.getMonthByDate(测试日期) as 月份 ,count(*) 检验批数 ,COUNT (case when 检验结果='NG' then 'NG'  else null end ) as NG批数 from   ";
            sql += "  (select d.deliveryid 接收单号,d.materialcode 料号,case max(TestFinalResult) when '拒收' then 'NG' else 'OK' end 检验结果,max(testtime) 测试日期  ";
            sql += "  from delivery d left join deliveryEMSOtherRec r on d.deliveryid=r.deliveryid    ";
            sql += "  right join (select receptid,Productcode,LotNo,max(testqty) testqty,max(TestFinalResult) TestFinalResult,min(testtime) testtime from IQC_TestList   ";
            sql += "  where  testtime >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "'  and testtime <= '" + txtenddate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and receptid is not null  ";
            sql += "  group by receptid,Productcode,Lotno) x on d.deliveryid=x.receptid and d.materialcode=x.Productcode and d.lotno=x.LotNo   ";
            sql += "  where  d.materialcode like '7%' and (" + sqlcustomer + ") group by d.deliveryid,d.materialcode  ";
            sql += "  ) z group by dbo.getMonthByDate(测试日期) ) t  inner join  (  ";
            sql += "  select dbo.getMonthByDate(来料日期) as 月份 ,count(HYT编码) 来料批数  from  ( ";
            sql += "  select  r.deliveryid 单号,m.materialcode  HYT编码, m.materialname 物料描述,r.vendorname 客户 ,r.arrivatedate 来料日期 ";
            sql += "  from deliveryEMSOtherRec r left join OEM_EMSHYTCusRelation cr on r.vendorcode=cr.cuscode and r.customer=cr.customer  ";
            sql += "  left join (select deliveryid,materialcode,min(transactiondate) eventtime  from delivery group by deliveryid,materialcode) d on cr.hytcode=d.materialcode and r.deliveryid=d.deliveryid ";
            sql += "  left join MaterialSpec m on cr.hytcode=m.materialcode  ";
            sql += "  where r.arrivatedate>='" + txtstartdate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and r.arrivatedate<= '" + txtenddate.DateTime.ToString("yyyy-MM-dd HH:mm:ss") + "' and (" + sqlcustomer + ") ";
            sql += "   ) la  group by dbo.getMonthByDate(来料日期) ) l on t.月份 = l.月份  order by  cast (t.月份 as int) asc  ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;

        }

        private void sBtnreset_Click(object sender, EventArgs e)
        {
 
            txtgoalrate.Text = "90";
            txtcomelotqty.Text = "";
            txttestqty.Text = "";
            txtNGqty.Text = "";
            txttestPassrate.Text = "";
            checkSHL.Checked = false;
            checkHCL.Checked = false;

            checkcustomer.DataSource = null;
            gridControl.DataSource = null;
            IQCreportchart.DataSource = null;


        }


        public void ExportToExcel(string title, params IPrintable[] panels)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = title;
            saveFileDialog.Title = "导出Excel";
            saveFileDialog.Filter = "Excel文件(*.xlsx)|*.xlsx|Excel文件(*.xls)|*.xls";
            DialogResult dialogResult = saveFileDialog.ShowDialog();
            if (dialogResult == DialogResult.Cancel)
                return;
            string FileName = saveFileDialog.FileName;
            PrintingSystem ps = new PrintingSystem();
            CompositeLink link = new CompositeLink(ps);
            ps.Links.Add(link);
            foreach (IPrintable panel in panels)
            {
                link.Links.Add(CreatePrintableLink(panel));
            }
            link.Landscape = true;  //横向           
            //判断是否有标题，有则设置         
            try
            {
                int count = 1;
                //在重复名称后加（序号）
                while (File.Exists(FileName))
                {
                    if (FileName.Contains(")."))
                    {
                        int start = FileName.LastIndexOf("(");
                        int end = FileName.LastIndexOf(").") - FileName.LastIndexOf("(") + 2;
                        FileName = FileName.Replace(FileName.Substring(start, end), string.Format("({0}).", count));
                    }
                    else
                    {
                        FileName = FileName.Replace(".", string.Format("({0}).", count));
                    }
                    count++;
                }
                if (FileName.LastIndexOf(".xlsx") >= FileName.Length - 5)
                {
                    XlsxExportOptions options = new XlsxExportOptions();
                    link.ExportToXlsx(FileName, options);
                }
                else
                {
                    XlsExportOptions options = new XlsExportOptions();
                    link.ExportToXls(FileName, options);
                }
                if (DevExpress.XtraEditors.XtraMessageBox.Show("保存成功，是否打开文件？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    System.Diagnostics.Process.Start(FileName);  

                progressBarControl1.Position = 0;
            }
            catch (Exception ex)
            {
                DevExpress.XtraEditors.XtraMessageBox.Show(ex.Message);
            }
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

        PrintableComponentLink CreatePrintableLink(IPrintable printable)
        {
            ChartControl chart = printable as ChartControl;
            if (chart != null)
                chart.OptionsPrint.SizeMode = DevExpress.XtraCharts.Printing.PrintSizeMode.Stretch;
            PrintableComponentLink printableLink = new PrintableComponentLink()
            {
                Component = printable
            };
            return printableLink;
        }



        private void sBtnreport_Click(object sender, EventArgs e)
        {
            DataTable dt = gridControl.DataSource as DataTable;
            if (dt == null || dt.Rows.Count < 1)
                return;
            if (gridView.RowCount < 1)
                return;
            string filename = IQCreportchart.Titles[0].ToString();          
            ExportToExcel(filename,gridControl, IQCreportchart);
        }


    }
}