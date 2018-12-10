using DevExpress.XtraCharts;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DX_QMS.Common;


namespace IQCReport
{
   
    public partial class OQCExceptTestShow : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public OQCExceptTestShow()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {
             
            this.comboBox1.Items.Add("EMS");
           
            ////string sql = @"select distinct badclass from OQC_baditem";
            ////DBHelper db = new DBHelper();
            //////控件名.DataSource=数据集.数据表
            ////comboBox3.DataSource = db.GetDataSet(sql).Tables[0];
            ////comboBox3.DisplayMember = "badclass";
            ////comboBox3.ValueMember = "badclass";
            bindbadclass();
            comboBox1.SelectedIndex = 0;
        }

        private void bindbadclass()
        {
            string sql = @"select distinct badclass from OQC_baditem";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            txtbadclass.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                if ((row["badclass"].ToString() != "其他") && (row["badclass"].ToString() != "功能不良"))
                txtbadclass.Items.Add(row["badclass"]);//
            }
        }
        private void bindngcustomer()
        {
            // string sql = @"select distinct customer from OQC_TestListNew where testresult='NG'";
            string sql = @"select distinct customer from OQC_TestListNew where badinformation <>''";
            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];
            comboBox2.Items.Clear();
            foreach (DataRow row in dt.Rows)
            {
                comboBox2.Items.Add(row["customer"]);//
            }
            comboBox2.Items.Add("所有客户");//
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                this.comboBox2.Text = "";
                this.comboBox2.Items.Clear();
                bindngcustomer();
                /*
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
                this.comboBox2.Items.Add("所有客户");*/
            }
            if (comboBox1.SelectedIndex==1)
            {
                this.comboBox2.Text = "";
                this.comboBox2.Items.Clear();
                this.comboBox2.Items.Add("主营");
            }
            comboBox2.SelectedIndex = 0;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            txtbadclass.SelectedIndex = -1;
            chartControl1.Series.Clear();
            gridView1.Columns.Clear();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            chartControl1.Series.Clear();
            gridView1.Columns.Clear();
            CreatTitleName();
            gridView1.OptionsView.ShowGroupPanel = true;
            gridView1.GroupPanelText = "OQC检验异常分析报表";
            DateTime starttime = txtarrivatedate1.Value.Date;
            DateTime endtime = txtarrivatedate2.Value.AddDays(1).Date;
            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("请输入业务类型");
                return;
            }
            if (comboBox2.SelectedItem==null)
            {
                MessageBox.Show("请输入客户");
                return;
            }
            if (comboBox1.SelectedItem=="EMS")
            {
                if (comboBox2.SelectedItem == "所有客户")
                    CreatOqcCustomer(starttime, endtime);
                else
                    CreatOqcType(starttime, endtime);
            }
        }
        private void CreatTitleName()
        {
            DevExpress.XtraCharts.ChartTitle chartTitle = new DevExpress.XtraCharts.ChartTitle();
            chartTitle.Text = "Pareto Chart of 出货检验报表";//标题内容
            chartTitle.TextColor = Color.Black; //颜色设置
            chartTitle.Font = new Font("Tahoma", 13);//字体类型字号
            chartTitle.Alignment = StringAlignment.Center;
            chartControl1.Titles.Clear();
            chartControl1.Titles.Add(chartTitle);
        }

        private void CreatOqcCustomer(DateTime starttime, DateTime endtime)
        {
            if (txtbadclass.SelectedItem==null)
            {
                int n2 = 0;
                int n3 = 0;
                string sql = @"select distinct badclass from OQC_baditem";
                DataTable dt0 = DbAccess.SelectBySql(sql).Tables[0];
                string[] NgType = new string[dt0.Rows.Count];
                for (int i = 0; i < dt0.Rows.Count; i++)
                {
                    NgType[i] = dt0.Rows[i]["badclass"].ToString();
                }
                int[] Number = new int[dt0.Rows.Count];
                DataTable dt = new DataTable();
                DataColumn dc1 = new DataColumn("不良类别", Type.GetType("System.String"));
                DataColumn dc2 = new DataColumn("不良批数", Type.GetType("System.Int32"));
                dt.Columns.Add(dc1);
                dt.Columns.Add(dc2);


                DataTable dt1 = new DataTable();
                DataColumn dc3 = new DataColumn("不良类别", Type.GetType("System.String"));
                DataColumn dc6 = new DataColumn("不良批数", Type.GetType("System.Int32"));//0305新添加不良批数。
                DataColumn dc5 = new DataColumn("Percent", Type.GetType("System.String"));//0305新添加不良批数对应的百分比。
                DataColumn dc4 = new DataColumn("Cum%", Type.GetType("System.String"));
                dt1.Columns.Add(dc3);
                dt1.Columns.Add(dc6);
                dt1.Columns.Add(dc5);
                dt1.Columns.Add(dc4);
                // string constr = "server=192.168.0.176;database=BarcodeNew;user id=sa;password=The0more7people0you7love3the7weaker8you8are;Pooling=false";
                //注意：此处连接了正式生产数据库。
                string constr = DbAccess.connSql;
                for (int i = 0; i < dt0.Rows.Count; i++)
                {
                    using (SqlConnection con = new SqlConnection(constr))
                    {
                        string sql1 = "QMS_NgTypeOqcforCustomer";
                        using (SqlCommand cmd = new SqlCommand(sql1, con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@ustarttime", starttime);
                            cmd.Parameters.AddWithValue("@uendtime", endtime);
                            cmd.Parameters.AddWithValue("@ngtype", NgType[i]);
                            con.Open();
                            Number[i] = (int)cmd.ExecuteScalar();
                        }
                    }
                    n2 += Number[i];
                    DataRow dr = dt.NewRow();
                    dr["不良类别"] = NgType[i];
                    dr["不良批数"] = Number[i];
                    if (Number[i] != 0)
                    {
                        dt.Rows.Add(dr);
                        // n1++;
                    }
                }
                dt.DefaultView.Sort = "不良批数 DESC";
                dt = dt.DefaultView.ToTable();

                int len = 0;
                len = dt.Rows.Count;
                int[] ArrayNumber = new int[len];
                for (int i = 0; i < len; i++)
                {
                    ArrayNumber[i] = Convert.ToInt32(dt.Rows[i]["不良批数"]);
                }

                for (int i = 0; i < len; i++)
                {
                    n3 += ArrayNumber[i];//dt.Rows[0][Column]
                    DataRow dr = dt1.NewRow();
                    dr["不良类别"] = dt.Rows[i]["不良类别"];
                    dr["不良批数"] = dt.Rows[i]["不良批数"];
                    dr["Percent"] = (ArrayNumber[i] / (double)n2).ToString("p");
                    dr["Cum%"] = (n3 / (double)n2).ToString("p");
                    dt1.Rows.Add(dr);
                }

                gridControl1.DataSource = dt1;
                this.gridView1.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                this.gridView1.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;


                chartControl1.Series.Clear();
                Series series1 = CreateSeries("不良批数", ViewType.Bar, dt1);
                chartControl1.Series.Add(series1);
                Series series3 = CreateSeries2(ViewType.Line, dt1);
                chartControl1.Series.Add(series3);

                //AxisRange.Auto
                AxisRange DIA = (AxisRange)((XYDiagram)chartControl1.Diagram).AxisY.Range;
                if (n2 > 0)
                {
                    DIA.Auto = false;
                    DIA.SetInternalMinMaxValues(0, n2);
                }
                XYDiagram diagram = (XYDiagram)chartControl1.Diagram;
                diagram.AxisY.Title.Visible = true;
                diagram.AxisY.Title.Alignment = StringAlignment.Center;
                diagram.AxisY.Title.Text = "不良批数";
                diagram.AxisY.Title.TextColor = Color.Black;
                diagram.AxisY.Title.Antialiasing = true;
                diagram.AxisY.Title.Font = new Font("Tahoma", 10);

                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Clear();
                SecondaryAxisY myAxisY = new SecondaryAxisY();
                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Add(myAxisY);
                ((LineSeriesView)series3.View).AxisY = myAxisY;
                myAxisY.Color = Color.Red;

                myAxisY.Range.Auto = false;
                myAxisY.Range.SetInternalMinMaxValues(0, 1);

                myAxisY.Title.Text = "不良占比";
                myAxisY.Title.Alignment = StringAlignment.Center;
                myAxisY.Title.Visible = true;
                myAxisY.Title.TextColor = Color.Black;
                myAxisY.Title.Font = new Font("Tahoma", 10);
                myAxisY.Title.Antialiasing = true;

                ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
                myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;//显示为百分数
            
             }
         else
            {
                int n2 = 0;
                int n3 = 0;  // where += " and hytcode = '" + PN + "' ";
                string badclasses = txtbadclass.SelectedItem.ToString();
                string sql = @"select distinct badphenomenon from OQC_baditem where badclass='"+badclasses+"'";
                DataTable dt0 = DbAccess.SelectBySql(sql).Tables[0];
                string[] NgType = new string[dt0.Rows.Count];
                string[] NgType1 = new string[dt0.Rows.Count];
                for (int i = 0; i < dt0.Rows.Count; i++)
                {
                    NgType[i] = dt0.Rows[i]["badphenomenon"].ToString();
                    NgType1[i] = badclasses + ";" + NgType[i];
                }
                int[] Number = new int[dt0.Rows.Count];
                DataTable dt = new DataTable();
                DataColumn dc1 = new DataColumn("不良现象", Type.GetType("System.String"));
                DataColumn dc2 = new DataColumn("不良批数", Type.GetType("System.Int32"));
                dt.Columns.Add(dc1);
                dt.Columns.Add(dc2);


                DataTable dt1 = new DataTable();
                DataColumn dc3 = new DataColumn("不良现象", Type.GetType("System.String"));
                DataColumn dc6 = new DataColumn("不良批数", Type.GetType("System.Int32"));//0305新添加不良批数。
                DataColumn dc5 = new DataColumn("Percent", Type.GetType("System.String"));//0305新添加不良批数对应的百分比。
                DataColumn dc4 = new DataColumn("Cum%", Type.GetType("System.String"));
                dt1.Columns.Add(dc3);
                dt1.Columns.Add(dc6);
                dt1.Columns.Add(dc5);
                dt1.Columns.Add(dc4);
                // string constr = "server=192.168.0.176;database=BarcodeNew;user id=sa;password=The0more7people0you7love3the7weaker8you8are;Pooling=false";
                //注意：此处连接了正式生产数据库。
                string constr = DbAccess.connSql;
                for (int i = 0; i < dt0.Rows.Count; i++)
                {
                    using (SqlConnection con = new SqlConnection(constr))
                    {
                        string sql1 = "QMS_NgTypeOqcforCustomer";
                        using (SqlCommand cmd = new SqlCommand(sql1, con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@ustarttime", starttime);
                            cmd.Parameters.AddWithValue("@uendtime", endtime);
                            cmd.Parameters.AddWithValue("@ngtype", NgType1[i]);
                            con.Open();
                            Number[i] = (int)cmd.ExecuteScalar();
                        }
                    }
                    n2 += Number[i];
                    DataRow dr = dt.NewRow();
                    dr["不良现象"] = NgType[i];
                    dr["不良批数"] = Number[i];
                    if (Number[i] != 0)
                    {
                        dt.Rows.Add(dr);
                        // n1++;
                    }
                }
                dt.DefaultView.Sort = "不良批数 DESC";
                dt = dt.DefaultView.ToTable();

                int len = 0;
                len = dt.Rows.Count;
                int[] ArrayNumber = new int[len];
                for (int i = 0; i < len; i++)
                {
                    ArrayNumber[i] = Convert.ToInt32(dt.Rows[i]["不良批数"]);
                }

                for (int i = 0; i < len; i++)
                {
                    n3 += ArrayNumber[i];//dt.Rows[0][Column]
                    DataRow dr = dt1.NewRow();
                    dr["不良现象"] = dt.Rows[i]["不良现象"];
                    dr["不良批数"] = dt.Rows[i]["不良批数"];
                    dr["Percent"] = (ArrayNumber[i] / (double)n2).ToString("p");
                    dr["Cum%"] = (n3 / (double)n2).ToString("p");
                    dt1.Rows.Add(dr);
                }
                gridControl1.DataSource = null;
                gridControl1.DataSource = dt1;
                this.gridView1.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                this.gridView1.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;


                chartControl1.Series.Clear();
                Series series1 = CreateSeries("不良批数", ViewType.Bar, dt1);
                chartControl1.Series.Add(series1);
                Series series3 = CreateSeries2(ViewType.Line, dt1);
                chartControl1.Series.Add(series3);

                //AxisRange.Auto
                AxisRange DIA = (AxisRange)((XYDiagram)chartControl1.Diagram).AxisY.Range;
                if (n2 > 0)
                {
                    DIA.Auto = false;
                    DIA.SetInternalMinMaxValues(0, n2);
                }
                XYDiagram diagram = (XYDiagram)chartControl1.Diagram;
                diagram.AxisY.Title.Visible = true;
                diagram.AxisY.Title.Alignment = StringAlignment.Center;
                diagram.AxisY.Title.Text = "不良批数";
                diagram.AxisY.Title.TextColor = Color.Black;
                diagram.AxisY.Title.Antialiasing = true;
                diagram.AxisY.Title.Font = new Font("Tahoma", 10);

                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Clear();
                SecondaryAxisY myAxisY = new SecondaryAxisY();
                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Add(myAxisY);
                ((LineSeriesView)series3.View).AxisY = myAxisY;
                myAxisY.Color = Color.Red;

                myAxisY.Range.Auto = false;
                myAxisY.Range.SetInternalMinMaxValues(0, 1);

                myAxisY.Title.Text = "不良占比";
                myAxisY.Title.Alignment = StringAlignment.Center;
                myAxisY.Title.Visible = true;
                myAxisY.Title.TextColor = Color.Black;
                myAxisY.Title.Font = new Font("Tahoma", 10);
                myAxisY.Title.Antialiasing = true;

                ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
                myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;//显示为百分数
        }
        }

        private void CreatOqcType(DateTime starttime, DateTime endtime)
        {
            if (txtbadclass.SelectedItem == null)
            {
                int n2 = 0;
                int n3 = 0;
                string sql = @"select distinct badclass from OQC_baditem";
                DataTable dt0 = DbAccess.SelectBySql(sql).Tables[0];
                string[] NgType = new string[dt0.Rows.Count];
                for (int i = 0; i < dt0.Rows.Count; i++)
                {
                    NgType[i] = dt0.Rows[i]["badclass"].ToString();
                }
                int[] Number = new int[dt0.Rows.Count];
                DataTable dt = new DataTable();
                DataColumn dc1 = new DataColumn("不良类别", Type.GetType("System.String"));
                DataColumn dc2 = new DataColumn("不良批数", Type.GetType("System.Int32"));
                dt.Columns.Add(dc1);
                dt.Columns.Add(dc2);


                DataTable dt1 = new DataTable();
                DataColumn dc3 = new DataColumn("不良类别", Type.GetType("System.String"));
                DataColumn dc6 = new DataColumn("不良批数", Type.GetType("System.Int32"));//0305新添加不良批数。
                DataColumn dc5 = new DataColumn("Percent", Type.GetType("System.String"));//0305新添加不良批数对应的百分比。
                DataColumn dc4 = new DataColumn("Cum%", Type.GetType("System.String"));
                dt1.Columns.Add(dc3);
                dt1.Columns.Add(dc6);
                dt1.Columns.Add(dc5);
                dt1.Columns.Add(dc4);
                // string constr = "server=192.168.0.176;database=BarcodeNew;user id=sa;password=The0more7people0you7love3the7weaker8you8are;Pooling=false";
                //注意：此处连接了正式生产数据库。
                string constr = DbAccess.connSql;
                for (int i = 0; i < dt0.Rows.Count; i++)
                {
                    using (SqlConnection con = new SqlConnection(constr))
                    {
                        string sql1 = "QMS_NgTypeOqcforall";
                        using (SqlCommand cmd = new SqlCommand(sql1, con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@ustarttime", starttime);
                            cmd.Parameters.AddWithValue("@uendtime", endtime);
                            cmd.Parameters.AddWithValue("@ngtype", NgType[i]);
                            cmd.Parameters.AddWithValue("@customer", comboBox2.SelectedItem.ToString());
                            con.Open();
                            Number[i] = (int)cmd.ExecuteScalar();
                        }
                    }
                    n2 += Number[i];
                    DataRow dr = dt.NewRow();
                    dr["不良类别"] = NgType[i];
                    dr["不良批数"] = Number[i];
                    if (Number[i] != 0)
                    {
                        dt.Rows.Add(dr);
                        // n1++;
                    }
                }
                dt.DefaultView.Sort = "不良批数 DESC";
                dt = dt.DefaultView.ToTable();

                int len = 0;
                len = dt.Rows.Count;
                int[] ArrayNumber = new int[len];
                for (int i = 0; i < len; i++)
                {
                    ArrayNumber[i] = Convert.ToInt32(dt.Rows[i]["不良批数"]);
                }

                for (int i = 0; i < len; i++)
                {
                    n3 += ArrayNumber[i];//dt.Rows[0][Column]
                    DataRow dr = dt1.NewRow();
                    dr["不良类别"] = dt.Rows[i]["不良类别"];
                    dr["不良批数"] = dt.Rows[i]["不良批数"];
                    dr["Percent"] = (ArrayNumber[i] / (double)n2).ToString("p");
                    dr["Cum%"] = (n3 / (double)n2).ToString("p");
                    dt1.Rows.Add(dr);
                }

                gridControl1.DataSource = dt1;
                this.gridView1.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                this.gridView1.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;


                chartControl1.Series.Clear();
                Series series1 = CreateSeries("不良批数", ViewType.Bar, dt1);
                chartControl1.Series.Add(series1);
                Series series3 = CreateSeries2(ViewType.Line, dt1);
                chartControl1.Series.Add(series3);

                //AxisRange.Auto
                AxisRange DIA = (AxisRange)((XYDiagram)chartControl1.Diagram).AxisY.Range;
                if (n2 > 0)
                {
                    DIA.Auto = false;
                    DIA.SetInternalMinMaxValues(0, n2);
                }
                XYDiagram diagram = (XYDiagram)chartControl1.Diagram;
                diagram.AxisY.Title.Visible = true;
                diagram.AxisY.Title.Alignment = StringAlignment.Center;
                diagram.AxisY.Title.Text = "不良批数";
                diagram.AxisY.Title.TextColor = Color.Black;
                diagram.AxisY.Title.Antialiasing = true;
                diagram.AxisY.Title.Font = new Font("Tahoma", 10);

                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Clear();
                SecondaryAxisY myAxisY = new SecondaryAxisY();
                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Add(myAxisY);
                ((LineSeriesView)series3.View).AxisY = myAxisY;
                myAxisY.Color = Color.Red;

                myAxisY.Range.Auto = false;
                myAxisY.Range.SetInternalMinMaxValues(0, 1);

                myAxisY.Title.Text = "不良占比";
                myAxisY.Title.Alignment = StringAlignment.Center;
                myAxisY.Title.Visible = true;
                myAxisY.Title.TextColor = Color.Black;
                myAxisY.Title.Font = new Font("Tahoma", 10);
                myAxisY.Title.Antialiasing = true;

                ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
                myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;//显示为百分数

            }
            else
            {
                int n2 = 0;
                int n3 = 0;  // where += " and hytcode = '" + PN + "' ";
                string badclasses = txtbadclass.SelectedItem.ToString();
                string sql = @"select distinct badphenomenon from OQC_baditem where badclass='" + badclasses + "'";
                DataTable dt0 = DbAccess.SelectBySql(sql).Tables[0];
                string[] NgType = new string[dt0.Rows.Count];
                string[] NgType1 = new string[dt0.Rows.Count];
                for (int i = 0; i < dt0.Rows.Count; i++)
                {
                    NgType[i] = dt0.Rows[i]["badphenomenon"].ToString();
                    NgType1[i] = badclasses + ";" + NgType[i];
                }
                int[] Number = new int[dt0.Rows.Count];
                DataTable dt = new DataTable();
                DataColumn dc1 = new DataColumn("不良现象", Type.GetType("System.String"));
                DataColumn dc2 = new DataColumn("不良批数", Type.GetType("System.Int32"));
                dt.Columns.Add(dc1);
                dt.Columns.Add(dc2);


                DataTable dt1 = new DataTable();
                DataColumn dc3 = new DataColumn("不良现象", Type.GetType("System.String"));
                DataColumn dc6 = new DataColumn("不良批数", Type.GetType("System.Int32"));//0305新添加不良批数。
                DataColumn dc5 = new DataColumn("Percent", Type.GetType("System.String"));//0305新添加不良批数对应的百分比。
                DataColumn dc4 = new DataColumn("Cum%", Type.GetType("System.String"));
                dt1.Columns.Add(dc3);
                dt1.Columns.Add(dc6);
                dt1.Columns.Add(dc5);
                dt1.Columns.Add(dc4);
                // string constr = "server=192.168.0.176;database=BarcodeNew;user id=sa;password=The0more7people0you7love3the7weaker8you8are;Pooling=false";
                //注意：此处连接了正式生产数据库。
                string constr = DbAccess.connSql;
                for (int i = 0; i < dt0.Rows.Count; i++)
                {
                    using (SqlConnection con = new SqlConnection(constr))
                    {
                        string sql1 = "QMS_NgTypeOqcforall";
                        using (SqlCommand cmd = new SqlCommand(sql1, con))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.AddWithValue("@ustarttime", starttime);
                            cmd.Parameters.AddWithValue("@uendtime", endtime);
                            cmd.Parameters.AddWithValue("@ngtype", NgType1[i]);
                            cmd.Parameters.AddWithValue("@customer", comboBox2.SelectedItem.ToString());
                            con.Open();
                            Number[i] = (int)cmd.ExecuteScalar();
                        }
                    }
                    n2 += Number[i];
                    DataRow dr = dt.NewRow();
                    dr["不良现象"] = NgType[i];
                    dr["不良批数"] = Number[i];
                    if (Number[i] != 0)
                    {
                        dt.Rows.Add(dr);
                        // n1++;
                    }
                }
                dt.DefaultView.Sort = "不良批数 DESC";
                dt = dt.DefaultView.ToTable();

                int len = 0;
                len = dt.Rows.Count;
                int[] ArrayNumber = new int[len];
                for (int i = 0; i < len; i++)
                {
                    ArrayNumber[i] = Convert.ToInt32(dt.Rows[i]["不良批数"]);
                }

                for (int i = 0; i < len; i++)
                {
                    n3 += ArrayNumber[i];//dt.Rows[0][Column]
                    DataRow dr = dt1.NewRow();
                    dr["不良现象"] = dt.Rows[i]["不良现象"];
                    dr["不良批数"] = dt.Rows[i]["不良批数"];
                    dr["Percent"] = (ArrayNumber[i] / (double)n2).ToString("p");
                    dr["Cum%"] = (n3 / (double)n2).ToString("p");
                    dt1.Rows.Add(dr);
                }

                gridControl1.DataSource = dt1;
                this.gridView1.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                this.gridView1.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;


                chartControl1.Series.Clear();
                Series series1 = CreateSeries("不良批数", ViewType.Bar, dt1);
                chartControl1.Series.Add(series1);
                Series series3 = CreateSeries2(ViewType.Line, dt1);
                chartControl1.Series.Add(series3);

                //AxisRange.Auto
                AxisRange DIA = (AxisRange)((XYDiagram)chartControl1.Diagram).AxisY.Range;
                if (n2 > 0)
                {
                    DIA.Auto = false;
                    DIA.SetInternalMinMaxValues(0, n2);
                }
                XYDiagram diagram = (XYDiagram)chartControl1.Diagram;
                diagram.AxisY.Title.Visible = true;
                diagram.AxisY.Title.Alignment = StringAlignment.Center;
                diagram.AxisY.Title.Text = "不良批数";
                diagram.AxisY.Title.TextColor = Color.Black;
                diagram.AxisY.Title.Antialiasing = true;
                diagram.AxisY.Title.Font = new Font("Tahoma", 10);

                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Clear();
                SecondaryAxisY myAxisY = new SecondaryAxisY();
                ((XYDiagram)chartControl1.Diagram).SecondaryAxesY.Add(myAxisY);
                ((LineSeriesView)series3.View).AxisY = myAxisY;
                myAxisY.Color = Color.Red;

                myAxisY.Range.Auto = false;
                myAxisY.Range.SetInternalMinMaxValues(0, 1);

                myAxisY.Title.Text = "不良占比";
                myAxisY.Title.Alignment = StringAlignment.Center;
                myAxisY.Title.Visible = true;
                myAxisY.Title.TextColor = Color.Black;
                myAxisY.Title.Font = new Font("Tahoma", 10);
                myAxisY.Title.Antialiasing = true;

                ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
                myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;//显示为百分数
            }

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
            series.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
            return series;
        }

        private Series CreateSeries2(ViewType line, DataTable dt1)
        {
            Series series = new Series("累计频率", ViewType.Line);
            int number1 = 0;
            foreach (DataRow dr in dt1.Rows)
            {
                string argument = Convert.ToString(dr[0]);
                string str1 = Convert.ToString(dr[3]).TrimEnd('%');
                //number1+= Convert.ToInt32(dr[1]);
                // double value = Convert.ToDouble(n3 / (double)n2);
                double value = Convert.ToDouble(str1) / 100;
                series.Points.Add(new SeriesPoint(argument, value));
            }
            series.ArgumentScaleType = ScaleType.Qualitative;
            series.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
            series.PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;
            return series;
        }

       
    }
}
