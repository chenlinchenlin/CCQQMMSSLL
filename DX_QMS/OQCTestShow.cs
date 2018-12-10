using DevExpress.Charts.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraCharts;
using DX_QMS.Common;

namespace IQCReport
{
    public partial class OQCTestShow : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        bool Judge;
        //public bool Judge { get; private set; }
        public object TitleName { get; private set; }

        public OQCTestShow()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.comboBox1.Items.Add("EMS");
            this.comboBox2.Items.Add("C1");
            this.comboBox2.Items.Add("C2");
            this.comboBox2.Items.Add("CZ");
            this.comboBox2.Items.Add("D1");
            this.comboBox2.Items.Add("E1");
            this.comboBox2.Items.Add("F1");
            this.comboBox2.Items.Add("G2");
            this.comboBox2.Items.Add("GE");
            this.comboBox2.Items.Add("GW");
            this.comboBox2.Items.Add("GX");
            this.comboBox2.Items.Add("HBTG");
            this.comboBox2.Items.Add("N2");
            this.comboBox2.Items.Add("Nufront");
            this.comboBox2.Items.Add("QTJ");
            this.comboBox2.Items.Add("SAE");
            this.comboBox2.Items.Add("T1");
            this.comboBox2.Items.Add("YBX");
            this.comboBox2.Items.Add("ZY");
            this.comboBox2.Items.Add("诺赛特");
            this.comboBox2.Items.Add("迅特");
            this.comboBox2.Items.Add("亚派");
            this.comboBox2.Items.Add("中移物联");
            this.comboBox2.Items.Add("XFT");
            this.comboBox2.Items.Add("ALK");
            this.comboBox2.Items.Add("SP");
            //0308新添加客户，新添加9个新客户。
            this.comboBox2.Items.Add("XTJ");
            this.comboBox2.Items.Add("D1R");
            this.comboBox2.Items.Add("A1");
            this.comboBox2.Items.Add("N2");
            this.comboBox2.Items.Add("B1");
            this.comboBox2.Items.Add("U1");
            this.comboBox2.Items.Add("XT");
            this.comboBox2.Items.Add("X1");
            this.comboBox2.Items.Add("Sokon");
            this.comboBox2.Items.Add("所有");
            this.comboBox3.Items.Add("日报");
            this.comboBox3.Items.Add("周报");
            this.comboBox3.Items.Add("月报");
            this.comboBox3.Items.Add("季报");
            this.comboBox3.Items.Add("年报");
            //this.comboBox4.Items.Add("作业不良");
            //this.comboBox4.Items.Add("周转不良");
            //this.comboBox4.Items.Add("制程不良");
            //this.comboBox4.Items.Add("错漏反");
            //this.comboBox4.Enabled = false;
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
        }

        private void btnsearch_Click(object sender, EventArgs e)
        {
            chartControl1.Series.Clear();
            gridView1.Columns.Clear();
            Judge = true;
            CreatTitleName();
            if (Judge)
            {
                DateTime starttime = txtarrivatedate1.Value;
                DateTime endtime = txtarrivatedate2.Value;
                if (comboBox2.SelectedIndex!=34)//0301添加了三个客户，将22改为25。
                {
                    string reporttype= comboBox3.SelectedItem.ToString();
                    string customer = comboBox2.SelectedItem.ToString();
                    CreatType(customer,reporttype, starttime, endtime);
                }
                else
                {
                    CreatAllCustomer(starttime, endtime);
                }
            }
        }

        private void CreatAllCustomer(DateTime starttime, DateTime endtime)
        {
            string constr = "server=192.168.0.176;database=BarcodeNew;user id=sa;password=The0more7people0you7love3the7weaker8you8are;Pooling=false";
            SqlConnection con = new SqlConnection(constr);
            con.Open();
            string sql = "QMS_OQCforallcustomer";
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@ustarttime", starttime.Date);
            cmd.Parameters.AddWithValue("@uendtime", endtime.AddDays(1).Date);
            gridView1.Columns.Clear();
            DataTable dt = new DataTable();
            da.Fill(dt);
            gridControl1.DataSource = dt;
            gridView1.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            gridView1.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            con.Close();
            //创建图表
            DevExpress.XtraCharts.ChartTitle chartTitle = new DevExpress.XtraCharts.ChartTitle();
           // ChartTitle chartTitle = new ChartTitle();
            chartTitle.Text = starttime.ToShortDateString().ToString() + "至" + endtime.ToShortDateString().ToString() + " EMS各客户出货检验报表";//标题内容
            chartTitle.TextColor = System.Drawing.Color.Black; //颜色设置
            chartTitle.Font = new Font("Tahoma", 12);//字体类型字号
            chartTitle.Alignment = StringAlignment.Center;
            chartControl1.Titles.Clear();
            chartControl1.Titles.Add(chartTitle);
            chartControl1.Series.Clear();
            //实例化图表元素
            Series series1 = CreateSeries("检验批数", ViewType.Bar, dt);
            chartControl1.Series.Add(series1);
            Series series2 = CreateSeries1("不良批数", ViewType.Bar, dt);
            chartControl1.Series.Add(series2);
            Series series3 = CreateSeries4(ViewType.Line, dt);
            chartControl1.Series.Add(series3);
            //series3.View.Color = Color.Black;

            //series3.PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;//series3显示百分号
            ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Clear();
            SecondaryAxisY myAxisY = new SecondaryAxisY();
            ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Add(myAxisY);
            ((LineSeriesView)series3.View).AxisY = myAxisY;
            myAxisY.Color = Color.Red;
            ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
            myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;//显示为百分数
            double max = 1;
            double min = 1;
            foreach (DataRow dr in dt.Rows)
            {
                string str1 = Convert.ToString(dr[3]).TrimEnd('%');
                double value = Convert.ToDouble(str1) / 100;
                if (value < min)
                {
                    min = value;
                }
            }
            if (min > 0.1)
            { min = min - 0.06; }
            myAxisY.Range.MinValue = min;
            myAxisY.Range.MaxValue = max;
            ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
            Series series4 = CreateSeries5(ViewType.Line, dt);
            chartControl1.Series.Add(series4);
            ((LineSeriesView)series4.View).AxisY = myAxisY;
           // series4.View.Color = Color.Red;//将目标批合格率的线条改为红色。
        }

        private void CreatType(string customer, string reporttype, DateTime starttime, DateTime endtime)
        {
            int n1 = 0;
            int n2 = 0;
            DataTable dt = new DataTable();
            if (reporttype == "周报")
            {
                if (starttime.DayOfWeek == 0)
                {
                    starttime = starttime.AddDays(-1);
                }
                if (endtime.DayOfWeek == 0)
                {
                    endtime = endtime.AddDays(-1);
                }

                DateTime startWeek = starttime.AddDays(1 - Convert.ToInt32(starttime.DayOfWeek.ToString("d"))).Date;//开始时间周一零时零分
                DateTime endWeek = endtime.AddDays(1 - Convert.ToInt32(endtime.DayOfWeek.ToString("d"))).Date.AddDays(7);//结束时间周一的时间
                TimeSpan ts = endWeek - startWeek;
                int to = ts.Days;
                int weeks = to / 7;//产生多少周，即产生多少个列。
                DataColumn dc1 = new DataColumn("日期", Type.GetType("System.String"));
                DataColumn dc2 = new DataColumn("检验批数", Type.GetType("System.Int32"));
                DataColumn dc3 = new DataColumn("不良批数", Type.GetType("System.Int32"));
                DataColumn dc4 = new DataColumn("目标批合格率", Type.GetType("System.String"));
                DataColumn dc5 = new DataColumn("批合格率", Type.GetType("System.String"));
                dt.Columns.Add(dc1);
                dt.Columns.Add(dc2);
                dt.Columns.Add(dc3);
                dt.Columns.Add(dc4);
                dt.Columns.Add(dc5);
                for (int i = 0; i < weeks; i++)
                {
                    string constr = "server=192.168.0.176;database=BarcodeNew;user id=sa;password=The0more7people0you7love3the7weaker8you8are;Pooling=false";
                    using (SqlConnection con = new SqlConnection(constr))
                    {
                        string sql1 = "QMS_OQCforall";
                        string sql2 = "QMS_OQCforng";
                        using (SqlCommand cmd = new SqlCommand(sql1, con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@ucustomer", comboBox2.SelectedItem);
                            cmd.Parameters.AddWithValue("@ustarttime", startWeek.AddDays(7 * i));
                            cmd.Parameters.AddWithValue("@uendtime", startWeek.AddDays(7 + 7 * i));
                            con.Open();
                            n1 = (int)cmd.ExecuteScalar();
                        }
                        using (SqlCommand cmd = new SqlCommand(sql2, con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@ucustomer", comboBox2.SelectedItem);
                            cmd.Parameters.AddWithValue("@ustarttime", startWeek.AddDays(7 * i));
                            cmd.Parameters.AddWithValue("@uendtime", startWeek.AddDays(7 + 7 * i));
                            n2 = (int)cmd.ExecuteScalar();
                        }
                    }
                    DataRow dr = dt.NewRow();
                    dr["日期"] = "第" + getweekofyear(startWeek.AddDays(7 * i)) + "周";
                    dr["检验批数"] = n1;
                    dr["不良批数"] = n2;
                    dr["批合格率"] = n1 == 0 ? 0.ToString() : (1 - n2 / (double)n1).ToString("p");
                    DateTime timepoint = DateTime.Parse("2017-08-01");
                    if (comboBox2.SelectedItem == "HBTG" || comboBox2.SelectedItem == "迅特" || comboBox2.SelectedItem == "GE" || comboBox2.SelectedItem == "XFT" || comboBox2.SelectedItem == "GW" || comboBox2.SelectedItem == "亚派" || comboBox2.SelectedItem == "G2" || comboBox2.SelectedItem == "E1" || comboBox2.SelectedItem == "ALK" || comboBox2.SelectedItem == "SP")
                        dr["目标批合格率"] = 0.998.ToString("p");
                    else
                    {
                        if (DateTime.Compare(startWeek.AddDays(7 * i), timepoint) < 0)
                            dr["目标批合格率"] = 0.985.ToString("p");
                        else
                            dr["目标批合格率"] = 0.992.ToString("p");
                    }
                    dt.Rows.Add(dr);
                }
                //开始创建图表元素
                Series series1 = CreateSeries("检验批数", ViewType.Bar, dt);
                chartControl1.Series.Add(series1);
                Series series2 = CreateSeries1("不良批数", ViewType.Bar, dt);
                chartControl1.Series.Add(series2);
                Series series3 = CreateSeries2(ViewType.Line, dt);
                chartControl1.Series.Add(series3);
                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Clear();
                SecondaryAxisY myAxisY = new SecondaryAxisY();
                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Add(myAxisY);
                ((LineSeriesView)series3.View).AxisY = myAxisY;
                myAxisY.Color = Color.Red;
                ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
                myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;//显示为百分数
                //myAxisY.Range.MinValue = 0.85;
                //myAxisY.Range.MaxValue = 1.00;
                double max = 1;
                double min = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    string str1 = Convert.ToString(dr[4]).TrimEnd('%');
                    double value = Convert.ToDouble(str1) / 100;
                    if (value < min)
                    {
                        min = value;
                    }
                }
                if (min > 0.1)
                { min = min - 0.06; }
                myAxisY.Range.MinValue = min;
                myAxisY.Range.MaxValue = max;
                ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
                Series series4 = CreateSeries3(ViewType.Line, dt);
                chartControl1.Series.Add(series4);
                ((LineSeriesView)series4.View).AxisY = myAxisY;
            }
            else if (reporttype == "月报")
            {
                DateTime startMonth = starttime.AddMonths(-1).Date.AddDays(1 - starttime.Day).AddMonths(1);//该月的1号零时零分
                DateTime endMonth = DateTime.Parse(endtime.AddDays(1 - endtime.Day).AddMonths(1).ToShortDateString());//下个月的1号零时零分
                int year = endMonth.Year - startMonth.Year;
                int months = year * 12 + (endMonth.Month - startMonth.Month);
                gridView1.Columns.Clear();
                DataColumn dc1 = new DataColumn("日期", Type.GetType("System.String"));
                DataColumn dc2 = new DataColumn("检验批数", Type.GetType("System.Int32"));
                DataColumn dc3 = new DataColumn("不良批数", Type.GetType("System.Int32"));
                DataColumn dc4 = new DataColumn("目标批合格率", Type.GetType("System.String"));
                DataColumn dc5 = new DataColumn("批合格率", Type.GetType("System.String"));
                dt.Columns.Add(dc1);
                dt.Columns.Add(dc2);
                dt.Columns.Add(dc3);
                dt.Columns.Add(dc4);
                dt.Columns.Add(dc5);
                for (int i = 0; i < months; i++)
                {
                    string constr = "server=192.168.0.176;database=BarcodeNew;user id=sa;password=The0more7people0you7love3the7weaker8you8are;Pooling=false";
                    using (SqlConnection con = new SqlConnection(constr))
                    {
                        //string sql1 = "QMS_searchtotal";//此处为总数的存储过程
                        //string sql2 = "QMS_searchNG";//此处为NG数的存储过程
                        ////string sql1 = "QMS_searchtotalnum";
                        ////string sql2 = "QMS_searchtotalnumng";
                        string sql1 = "QMS_OQCforall";
                        string sql2 = "QMS_OQCforng";
                        using (SqlCommand cmd = new SqlCommand(sql1, con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@ucustomer", comboBox2.SelectedItem);
                            cmd.Parameters.AddWithValue("@ustarttime", startMonth.AddMonths(i));
                            cmd.Parameters.AddWithValue("@uendtime", startMonth.AddMonths(i + 1));
                            con.Open();
                            n1 = (int)cmd.ExecuteScalar();
                        }
                        using (SqlCommand cmd = new SqlCommand(sql2, con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@ucustomer", comboBox2.SelectedItem);
                            cmd.Parameters.AddWithValue("@ustarttime", startMonth.AddMonths(i));
                            cmd.Parameters.AddWithValue("@uendtime", startMonth.AddMonths(i + 1).AddSeconds(-1));
                            n2 = (int)cmd.ExecuteScalar();
                        }
                    }
                    // dt.Columns.Add("{0}", startDay.ToString("yyyy-MM-dd"), typeof(DateTime));
                    DataRow dr = dt.NewRow();
                    dr["日期"] = startMonth.AddMonths(i).ToString("yyyy-MM");
                    dr["检验批数"] = n1;
                    dr["不良批数"] = n2;
                    // dr["批合格率"] = n1==0?0.ToString():(1 - n2 / (double)n1).ToString("p");
                    dr["批合格率"] = n1 == 0 ? 0.ToString() : (1 - n2 / (double)n1).ToString("p");

                    //if (startDay.AddDays(i) > '2017-07-01')
                    DateTime timepoint = DateTime.Parse("2017-08-01");
                    if (comboBox2.SelectedItem == "HBTG" || comboBox2.SelectedItem == "迅特" || comboBox2.SelectedItem == "GE" || comboBox2.SelectedItem == "XFT" || comboBox2.SelectedItem == "GW" || comboBox2.SelectedItem == "亚派" || comboBox2.SelectedItem == "G2" || comboBox2.SelectedItem == "E1" || comboBox2.SelectedItem == "ALK" || comboBox2.SelectedItem == "SP")
                        dr["目标批合格率"] = 0.998.ToString("p");
                    else
                    {
                        if (DateTime.Compare(startMonth.AddMonths(i), timepoint) < 0)
                            dr["目标批合格率"] = 0.985.ToString("p");
                        else
                            dr["目标批合格率"] = 0.992.ToString("p");
                    }
                    dt.Rows.Add(dr);
                }
                //开始创建图表元素
                chartControl1.Series.Clear();
                Series series1 = CreateSeries("检验批数", ViewType.Bar, dt);
                chartControl1.Series.Add(series1);
                Series series2 = CreateSeries1("不良批数", ViewType.Bar, dt);
                chartControl1.Series.Add(series2);
                Series series3 = CreateSeries2(ViewType.Line, dt);
                chartControl1.Series.Add(series3);
                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Clear();
                SecondaryAxisY myAxisY = new SecondaryAxisY();
                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Add(myAxisY);
                ((LineSeriesView)series3.View).AxisY = myAxisY;
                myAxisY.Color = Color.Red;
                ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
                myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;//显示为百分数
                ////myAxisY.Range.MinValue = 0.85;
                ////myAxisY.Range.MaxValue = 1.00;
                double max = 1;
                double min = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    string str1 = Convert.ToString(dr[4]).TrimEnd('%');
                    double value = Convert.ToDouble(str1) / 100;
                    if (value < min)
                    {
                        min = value;
                    }
                }
                if (min > 0.1)
                { min = min - 0.06; }
                myAxisY.Range.MinValue = min;
                myAxisY.Range.MaxValue = max;
                ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
                // myAxisY.Label.format
                Series series4 = CreateSeries3(ViewType.Line, dt);
                chartControl1.Series.Add(series4);
                ((LineSeriesView)series4.View).AxisY = myAxisY;
            }
            else if (reporttype == "季报")
            {
                DateTime startQuarter = starttime.AddMonths(0 - (starttime.Month - 1) % 3).AddDays(1 - starttime.Day).Date;//本季度第一天
                DateTime endQuarter = endtime.AddMonths(0 - (endtime.Month - 1) % 3).AddDays(1 - endtime.Day).Date.AddMonths(3);//取得结束时间这一个季度的最后一天
                int year = endQuarter.Year - startQuarter.Year;
                int months = year * 12 + (endQuarter.Month - startQuarter.Month);
                int ts = months / 3;//一共有几个季度
                gridView1.Columns.Clear();
                DataColumn dc1 = new DataColumn("日期", Type.GetType("System.String"));
                DataColumn dc2 = new DataColumn("检验批数", Type.GetType("System.Int32"));
                DataColumn dc3 = new DataColumn("不良批数", Type.GetType("System.Int32"));
                DataColumn dc4 = new DataColumn("目标批合格率", Type.GetType("System.String"));
                DataColumn dc5 = new DataColumn("批合格率", Type.GetType("System.String"));
                dt.Columns.Add(dc1);
                dt.Columns.Add(dc2);
                dt.Columns.Add(dc3);
                dt.Columns.Add(dc4);
                dt.Columns.Add(dc5);
                for (int i = 0; i < ts; i++)
                {
                    string constr = "server=192.168.0.176;database=BarcodeNew;user id=sa;password=The0more7people0you7love3the7weaker8you8are;Pooling=false";
                    using (SqlConnection con = new SqlConnection(constr))
                    {
                        //string sql1 = "QMS_searchtotal";//此处为总数的存储过程
                        //string sql2 = "QMS_searchNG";//此处为NG数的存储过程
                        ////string sql1 = "QMS_searchtotalnum";
                        ////string sql2 = "QMS_searchtotalnumng";
                        string sql1 = "QMS_OQCforall";
                        string sql2 = "QMS_OQCforng";
                        using (SqlCommand cmd = new SqlCommand(sql1, con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@ucustomer", comboBox2.SelectedItem);
                            cmd.Parameters.AddWithValue("@ustarttime", startQuarter.AddMonths(i * 3));
                            cmd.Parameters.AddWithValue("@uendtime", startQuarter.AddMonths(i * 3 + 3));
                            con.Open();
                            n1 = (int)cmd.ExecuteScalar();
                        }
                        using (SqlCommand cmd = new SqlCommand(sql2, con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@ucustomer", comboBox2.SelectedItem);
                            cmd.Parameters.AddWithValue("@ustarttime", startQuarter.AddMonths(i * 3));
                            cmd.Parameters.AddWithValue("@uendtime", startQuarter.AddMonths(i * 3 + 3));
                            n2 = (int)cmd.ExecuteScalar();
                        }
                    }
                    DataRow dr = dt.NewRow();
                    dr["日期"] = startQuarter.AddMonths(i * 3).ToString("yyyy-MM") + "至" + startQuarter.AddMonths(i * 3 + 2).ToString("yyyy-MM");
                    dr["检验批数"] = n1;
                    dr["不良批数"] = n2;
                    dr["批合格率"] = n1 == 0 ? 0.ToString() : (1 - n2 / (double)n1).ToString("p");
                    DateTime timepoint = DateTime.Parse("2017-08-01");
                    if (comboBox2.SelectedItem == "HBTG" || comboBox2.SelectedItem == "迅特" || comboBox2.SelectedItem == "GE" || comboBox2.SelectedItem == "XFT" || comboBox2.SelectedItem == "GW" || comboBox2.SelectedItem == "亚派" || comboBox2.SelectedItem == "G2" || comboBox2.SelectedItem == "E1" || comboBox2.SelectedItem == "ALK" || comboBox2.SelectedItem == "SP")
                        dr["目标批合格率"] = 0.998.ToString("p");
                    else
                    {
                        if (DateTime.Compare(startQuarter.AddMonths(3 * i), timepoint) < 0)
                            dr["目标批合格率"] = 0.985.ToString("p");
                        else
                            dr["目标批合格率"] = 0.992.ToString("p");
                    }
                    dt.Rows.Add(dr);
                }
                //开始创建图表元素
                chartControl1.Series.Clear();
                Series series1 = CreateSeries("检验批数", ViewType.Bar, dt);
                chartControl1.Series.Add(series1);
                Series series2 = CreateSeries1("不良批数", ViewType.Bar, dt);
                chartControl1.Series.Add(series2);
                Series series3 = CreateSeries2(ViewType.Line, dt);
                chartControl1.Series.Add(series3);
                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Clear();
                SecondaryAxisY myAxisY = new SecondaryAxisY();
                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Add(myAxisY);
                ((LineSeriesView)series3.View).AxisY = myAxisY;
                myAxisY.Color = Color.Red;
                ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
                myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;//显示为百分数
                // myAxisY.Range.MinValue = 0.85;
                // myAxisY.Range.MaxValue = 1.00;
                double max = 1;
                double min = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    string str1 = Convert.ToString(dr[4]).TrimEnd('%');
                    double value = Convert.ToDouble(str1) / 100;
                    if (value < min)
                    {
                        min = value;
                    }
                }
                if (min > 0.1)
                { min = min - 0.06; }
                myAxisY.Range.MinValue = min;
                myAxisY.Range.MaxValue = max;
                ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
                // myAxisY.Label.format
                Series series4 = CreateSeries3(ViewType.Line, dt);
                chartControl1.Series.Add(series4);
                ((LineSeriesView)series4.View).AxisY = myAxisY;
            }
            else if (reporttype == "年报")
            {
                DateTime startYear = new DateTime(starttime.Year, 1, 1);  //本年年初  
                                                                          //  DateTime endYear = new DateTime(starttime.Year, 12, 31);  //本年年末  
                gridView1.Columns.Clear();
                DataColumn dc1 = new DataColumn("日期", Type.GetType("System.String"));
                DataColumn dc2 = new DataColumn("检验批数", Type.GetType("System.Int32"));
                DataColumn dc3 = new DataColumn("不良批数", Type.GetType("System.Int32"));
                DataColumn dc4 = new DataColumn("目标批合格率", Type.GetType("System.String"));
                DataColumn dc5 = new DataColumn("批合格率", Type.GetType("System.String"));
                dt.Columns.Add(dc1);
                dt.Columns.Add(dc2);
                dt.Columns.Add(dc3);
                dt.Columns.Add(dc4);
                dt.Columns.Add(dc5);
                for (int i = 0; i < 12; i++)
                {
                    string constr = "server=192.168.0.176;database=BarcodeNew;user id=sa;password=The0more7people0you7love3the7weaker8you8are;Pooling=false";
                    using (SqlConnection con = new SqlConnection(constr))
                    {
                        string sql1 = "QMS_OQCforall";
                        string sql2 = "QMS_OQCforng";
                        using (SqlCommand cmd = new SqlCommand(sql1, con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@ucustomer", comboBox2.SelectedItem);
                            cmd.Parameters.AddWithValue("@ustarttime", startYear.AddMonths(i));
                            cmd.Parameters.AddWithValue("@uendtime", startYear.AddMonths(i + 1));
                            con.Open();
                            n1 = (int)cmd.ExecuteScalar();
                        }
                        using (SqlCommand cmd = new SqlCommand(sql2, con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@ucustomer", comboBox2.SelectedItem);
                            cmd.Parameters.AddWithValue("@ustarttime", startYear.AddMonths(i));
                            cmd.Parameters.AddWithValue("@uendtime", startYear.AddMonths(i + 1));
                            n2 = (int)cmd.ExecuteScalar();
                        }
                    }
                    DataRow dr = dt.NewRow();
                    dr["日期"] = startYear.Year.ToString() + "年" + startYear.AddMonths(i).Month.ToString() + "月";
                    dr["检验批数"] = n1;
                    dr["不良批数"] = n2;
                    dr["批合格率"] = n1 == 0 ? 0.ToString() : (1 - n2 / (double)n1).ToString("p");
                    DateTime timepoint = DateTime.Parse("2017-08-01");
                    if (comboBox2.SelectedItem == "HBTG" || comboBox2.SelectedItem == "迅特" || comboBox2.SelectedItem == "GE" || comboBox2.SelectedItem == "XFT" || comboBox2.SelectedItem == "GW" || comboBox2.SelectedItem == "亚派" || comboBox2.SelectedItem == "G2" || comboBox2.SelectedItem == "E1" || comboBox2.SelectedItem == "ALK" || comboBox2.SelectedItem == "SP")
                        dr["目标批合格率"] = 0.998.ToString("p");
                    else
                    {
                        if (DateTime.Compare(startYear.AddMonths(i), timepoint) < 0)
                            dr["目标批合格率"] = 0.985.ToString("p");
                        else
                            dr["目标批合格率"] = 0.992.ToString("p");
                    }
                    dt.Rows.Add(dr);
                }
                //开始创建图表元素
                chartControl1.Series.Clear();
                Series series1 = CreateSeries("检验批数", ViewType.Bar, dt);
                chartControl1.Series.Add(series1);
                Series series2 = CreateSeries1("不良批数", ViewType.Bar, dt);
                chartControl1.Series.Add(series2);

                Series series3 = CreateSeries2(ViewType.Line, dt);
                chartControl1.Series.Add(series3);
                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Clear();
                SecondaryAxisY myAxisY = new SecondaryAxisY();
                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Add(myAxisY);
                ((LineSeriesView)series3.View).AxisY = myAxisY;
                myAxisY.Color = Color.Red;
                ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
                myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;//显示为百分数
                //myAxisY.Range.MinValue = 0.85;
                //myAxisY.Range.MaxValue = 1.00;
                double max = 1;
                double min = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    string str1 = Convert.ToString(dr[4]).TrimEnd('%');
                    double value = Convert.ToDouble(str1) / 100;
                    if (value < min)
                    {
                        min = value;
                    }
                }
                if (min > 0.1)
                { min = min - 0.06; }
                myAxisY.Range.MinValue = min;
                myAxisY.Range.MaxValue = max;
                ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
                Series series4 = CreateSeries3(ViewType.Line, dt);
                chartControl1.Series.Add(series4);
                ((LineSeriesView)series4.View).AxisY = myAxisY;

            }
            else if (reporttype == "日报")
            {
                DateTime startDay = starttime.Date;
                DateTime endDay = endtime.AddDays(1).Date;
                TimeSpan ts = endDay - startDay;
                int days = ts.Days;
                //   DataTable dt = new DataTable();
                //  DataSet ds = new DataSet();
                // dt.Columns.Add("日期", typeof(string));
                // DataRow dr1 = new DataRow("日期", Type.GetType(System.String));
                gridView1.Columns.Clear();
                DataColumn dc1 = new DataColumn("日期", Type.GetType("System.String"));
                DataColumn dc2 = new DataColumn("检验批数", Type.GetType("System.Int32"));
                DataColumn dc3 = new DataColumn("不良批数", Type.GetType("System.Int32"));
                DataColumn dc4 = new DataColumn("目标批合格率", Type.GetType("System.String"));
                DataColumn dc5 = new DataColumn("批合格率", Type.GetType("System.String"));
                dt.Columns.Add(dc1);
                dt.Columns.Add(dc2);
                dt.Columns.Add(dc3);
                dt.Columns.Add(dc4);
                dt.Columns.Add(dc5);
                for (int i = 0; i < days; i++)
                {
                    string constr = "server=192.168.0.176;database=BarcodeNew;user id=sa;password=The0more7people0you7love3the7weaker8you8are;Pooling=false";
                    using (SqlConnection con = new SqlConnection(constr))
                    {
                        string sql1 = "QMS_OQCforall";
                        string sql2 = "QMS_OQCforng";
                        using (SqlCommand cmd = new SqlCommand(sql1, con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@ucustomer", comboBox2.SelectedItem);
                            cmd.Parameters.AddWithValue("@ustarttime", startDay.AddDays(i));
                            cmd.Parameters.AddWithValue("@uendtime", startDay.AddDays(i + 1));
                            con.Open();
                            n1 = (int)cmd.ExecuteScalar();
                        }
                        using (SqlCommand cmd = new SqlCommand(sql2, con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@ucustomer", comboBox2.SelectedItem);
                            cmd.Parameters.AddWithValue("@ustarttime", startDay.AddDays(i));
                            cmd.Parameters.AddWithValue("@uendtime", startDay.AddDays(i + 1));
                            n2 = (int)cmd.ExecuteScalar();
                        }
                    }
                    DataRow dr = dt.NewRow();
                    dr["日期"] = startDay.AddDays(i).ToString("yyyy-MM-dd");
                    dr["检验批数"] = n1;
                    dr["不良批数"] = n2;
                    dr["批合格率"] = n1 == 0 ? 0.ToString() : (1 - n2 / (double)n1).ToString("p");
                    DateTime timepoint = DateTime.Parse("2017-08-01");

                    if (comboBox2.SelectedItem == "HBTG" || comboBox2.SelectedItem == "迅特" || comboBox2.SelectedItem == "GE" || comboBox2.SelectedItem == "XFT" || comboBox2.SelectedItem == "GW" || comboBox2.SelectedItem == "亚派" || comboBox2.SelectedItem == "G2" || comboBox2.SelectedItem == "E1" || comboBox2.SelectedItem == "ALK" || comboBox2.SelectedItem == "SP")
                        dr["目标批合格率"] = 0.998.ToString("p");
                    else
                    {
                        if (DateTime.Compare(startDay.AddDays(i), timepoint) < 0)
                            dr["目标批合格率"] = 0.985.ToString("p");
                        else
                            dr["目标批合格率"] = 0.992.ToString("p");
                    }
                    dt.Rows.Add(dr);
                }
                //开始创建图表元素
                chartControl1.Series.Clear();
                Series series1 = CreateSeries("检验批数", ViewType.Bar, dt);
                chartControl1.Series.Add(series1);
                Series series2 = CreateSeries1("不良批数", ViewType.Bar, dt);
                chartControl1.Series.Add(series2);

                Series series3 = CreateSeries2(ViewType.Line, dt);
                chartControl1.Series.Add(series3);
                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Clear();
                SecondaryAxisY myAxisY = new SecondaryAxisY();
                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Add(myAxisY);
                ((LineSeriesView)series3.View).AxisY = myAxisY;
                myAxisY.Color = Color.Red;
                ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
                myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;//显示为百分数
                // myAxisY.Range.MinValue = 0.85;
                //myAxisY.Range.MaxValue = 1.00;
                double max = 1;
                double min = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    string str1 = Convert.ToString(dr[4]).TrimEnd('%');
                    double value = Convert.ToDouble(str1) / 100;
                    if (value < min)
                    {
                        min = value;
                    }
                }
                if (min > 0.1)
                { min = min - 0.06; }
                myAxisY.Range.MinValue = min;
                myAxisY.Range.MaxValue = max;
                ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
                Series series4 = CreateSeries3(ViewType.Line, dt);
                chartControl1.Series.Add(series4);
                ((LineSeriesView)series4.View).AxisY = myAxisY;
            }
            gridControl1.DataSource = dt;
            this.gridView1.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridView1.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

        }
        private Series CreateSeries(string caption, ViewType viewType, DataTable dt)
        {
            Series series = new Series(caption, viewType);

            foreach (DataRow dr in dt.Rows)
            {
                string argument = Convert.ToString(dr[0]);
                double value = Convert.ToDouble(dr[1]);
                series.Points.Add(new SeriesPoint(argument, value));
            }
            series.ArgumentScaleType = ScaleType.Qualitative;
            series.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;//显示标注标签
            return series;
        }
        private Series CreateSeries1(string caption, ViewType viewType, DataTable dt)
        {
            Series series = new Series(caption, viewType);
            foreach (DataRow dr in dt.Rows)
            {
                string argument = Convert.ToString(dr[0]);
                double value = Convert.ToDouble(dr[2]);
                series.Points.Add(new SeriesPoint(argument, value));
            }
            series.ArgumentScaleType = ScaleType.Qualitative;
            series.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;//显示标注标签
            return series;
        }
        private Series CreateSeries2(ViewType viewType, DataTable dt)
        {
            Series series = new Series("目标批合格率", viewType);
            foreach (DataRow dr in dt.Rows)
            {
                string argument = Convert.ToString(dr[0]);
                string str1 = "";
                str1 = Convert.ToString(dr[3]).TrimEnd('%');
                double value = Convert.ToDouble(str1) / 100;
                series.Points.Add(new SeriesPoint(argument, value));
            }
            series.ArgumentScaleType = ScaleType.Qualitative;
            series.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;//显示标注标签
            series.PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;//series3显示百分号
            return series;
        }
        private Series CreateSeries3(ViewType viewType, DataTable dt)
        {
            Series series = new Series("批合格率", viewType);

            foreach (DataRow dr in dt.Rows)
            {
                string argument = Convert.ToString(dr[0]);
                string str1 = "";
                str1 = Convert.ToString(dr[4]).TrimEnd('%');
                double value = Convert.ToDouble(str1) / 100;
                series.Points.Add(new SeriesPoint(argument, value));
            }
            series.ArgumentScaleType = ScaleType.Qualitative;
            series.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;//显示标注标签
            series.PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;//series3显示百分号
            return series;
        }
        private Series CreateSeries4(ViewType viewType, DataTable dt)
        {
            Series series = new Series("批合格率", viewType);

            foreach (DataRow dr in dt.Rows)
            {
                string argument = Convert.ToString(dr[0]);
                string str1 = "";
                str1 = Convert.ToString(dr[3]).TrimEnd('%');
                double value = Convert.ToDouble(str1) / 100;
                series.Points.Add(new SeriesPoint(argument, value));
            }
            series.PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;//series显示百分号
            series.ArgumentScaleType = ScaleType.Qualitative;
            series.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;//显示标注标签
            return series;
        }
        private Series CreateSeries5(ViewType viewType, DataTable dt)
        {
            Series series = new Series("目标批合格率", viewType);
            foreach (DataRow dr in dt.Rows)
            {
                string argument = Convert.ToString(dr[0]);
                string str1 = "";
                str1 = Convert.ToString(dr[4]).TrimEnd('%');
                double value = Convert.ToDouble(str1) / 100;
                series.Points.Add(new SeriesPoint(argument, value));
            }
            series.ArgumentScaleType = ScaleType.Qualitative;
            series.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;//显示标注标签
            series.PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;//series3显示百分号
           
            return series;
        }

        private int getweekofyear(DateTime dt)
        {
            GregorianCalendar gc = new GregorianCalendar();
            int weekOfYear = gc.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
            return weekOfYear;
        }

        private void CreatTitleName()
        {
            string title1 = "";
            string title2 = "";
            string title3 = "";
            string title4 = "";
            string TitleName = "";
            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("请输入业务类型");
                Judge = false;
                return;
            }
            if (comboBox2.SelectedItem==null)
            {
                MessageBox.Show("请选择客户");
                Judge = false;
                return;
            }
            if (comboBox2.SelectedIndex!=34)
            {
                if (comboBox3.SelectedItem == null)
                {
                    MessageBox.Show("请选择报表类型");
                    Judge = false;
                    return;
                }
                if (comboBox1 .SelectedIndex==0)
            {
                title1 = "EMS ";
            }
            if(comboBox2.SelectedIndex != 34)
            {
                title2 = comboBox2.SelectedItem.ToString();
                title3 = "客户" + "出货检验";
                title4 = comboBox3.SelectedItem.ToString();
                TitleName = title1 + title2 + title3 + title4;
                gridView1.OptionsView.ShowGroupPanel = true;
                gridView1.GroupPanelText = TitleName.ToString();
            }
            }
            
          //  chartControl1.Series.Clear();
            DevExpress.XtraCharts.ChartTitle chartTitle = new DevExpress.XtraCharts.ChartTitle();
            chartTitle.Text = TitleName;//标题内容
            chartTitle.TextColor = Color.Black; //颜色设置
            chartTitle.Font = new Font("Tahoma", 12);//字体类型字号
            chartTitle.Alignment = StringAlignment.Center;
            chartControl1.Titles.Clear();
            chartControl1.Titles.Add(chartTitle);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == 34)
            {
                //comboBox4.SelectedItem = null;
                comboBox3.SelectedItem = null;
                comboBox3.Enabled = false;
               // comboBox4.Enabled = false;
                TitleName = "EMS所有客户出货检验报表";
                gridView1.OptionsView.ShowGroupPanel = true;
                gridView1.GroupPanelText = TitleName.ToString();
            }
            else
            {
               // comboBox4.Enabled = true;
                comboBox3.Enabled = true;
            }

        }

        private void gridControl1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //SaveFileDialog filedialog = new SaveFileDialog();
            ////filedialog.Title = "导出jpeg";
            //filedialog.Filter = "Image files(*.jpeg,*.gif)|*.jpeg;*.gif";
            //filedialog.FilterIndex = 0;
            //filedialog.RestoreDirectory = true;
            //DialogResult dialogresult = filedialog.ShowDialog(this);
            //if (dialogresult == DialogResult.OK)
            //{
            //    DevExpress.xtraprinting.xlsexportoptions options = new DevExpress.xtraprinting.xlsexportoptions();
            //    chartControl1.ExportToImage(filedialog.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
            //    DevExpress.XtraEditors.XtraMessageBox.Show("保存成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "查询报表";
            saveFileDialog.Filter = "Jpeg文件(*.jpg)|*.jpg";
            DialogResult dialogResult = saveFileDialog.ShowDialog(this);
            if (dialogResult == DialogResult.OK)
            {
                chartControl1.ExportToImage(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Jpeg);
                DevExpress.XtraEditors.XtraMessageBox.Show("保存成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.Title = "导出Excel";
            fileDialog.Filter = "Excel文件(*.xls)|*.xls";
            DialogResult dialogResult = fileDialog.ShowDialog(this);
            if (dialogResult == DialogResult.OK)
            {
                DevExpress.XtraPrinting.XlsExportOptions options = new DevExpress.XtraPrinting.XlsExportOptions();
                gridControl1.ExportToXls(fileDialog.FileName);
                DevExpress.XtraEditors.XtraMessageBox.Show("保存成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
          //  comboBox4.SelectedIndex = -1;
        }
    }
}
