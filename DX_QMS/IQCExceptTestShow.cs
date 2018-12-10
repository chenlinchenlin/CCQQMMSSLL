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

namespace IQCReport
{
    public partial class IQCExceptTestShow : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public IQCExceptTestShow()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            this.comboBox1.Items.Add("主营(结构+电子)");
            this.comboBox1.Items.Add("主营结构料");
            this.comboBox1.Items.Add("主营电子料");
            this.comboBox1.Items.Add("EMS");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox3.SelectedItem = null;
            if (comboBox1.SelectedIndex == 3)
            {
                this.comboBox2.Text = "";
                this.comboBox2.Items.Clear();

                this.comboBox2.Items.Add("A1");
                this.comboBox2.Items.Add("ANKE");
                this.comboBox2.Items.Add("C1");
                this.comboBox2.Items.Add("C2");
                this.comboBox2.Items.Add("CZ");
                this.comboBox2.Items.Add("D1");
                this.comboBox2.Items.Add("D1R");
                this.comboBox2.Items.Add("Finsar");
                this.comboBox2.Items.Add("GE");
                this.comboBox2.Items.Add("GX");
                this.comboBox2.Items.Add("HBTG");
                this.comboBox2.Items.Add("N2");
                this.comboBox2.Items.Add("QTJ");
                this.comboBox2.Items.Add("SAE");
                this.comboBox2.Items.Add("T1");
                this.comboBox2.Items.Add("TR");
                this.comboBox2.Items.Add("XT");
                this.comboBox2.Items.Add("YBX");
                this.comboBox2.Items.Add("ZY");
                this.comboBox2.Items.Add("所有客户");


                this.comboBox3.Items.Add("电子料");
                this.comboBox3.Items.Add("结构件");
                this.comboBox3.Items.Add("包材");
                this.comboBox3.Items.Add("辅料");
                this.comboBox3.Items.Add("PCB/FCB");
                this.comboBox3.Items.Add("电容");
                this.comboBox3.Items.Add("电阻");
                this.comboBox3.Items.Add("电感");
                this.comboBox3.Items.Add("半导体器件");
                this.comboBox3.Items.Add("IC");
                this.comboBox3.Items.Add("晶振(石英)");
                this.comboBox3.Items.Add("开关类");
                this.comboBox3.Items.Add("连接器");
                this.comboBox3.Items.Add("小五金类");
                this.comboBox3.Items.Add("线材");
               // this.comboBox3.Items.Add("所有物料");
            }
            else if (comboBox1.SelectedIndex == 0)
            {

                this.comboBox2.Text = "";
                this.comboBox2.Items.Clear();
                this.comboBox3.Items.Clear();
                this.comboBox2.Items.Add("主营");

                this.comboBox3.Items.Add("配件原材料");
                this.comboBox3.Items.Add("电子料");
                this.comboBox3.Items.Add("结构件");
                this.comboBox3.Items.Add("包材");
                this.comboBox3.Items.Add("辅料");
                this.comboBox3.Items.Add("外购成品");
                //this.comboBox3.Items.Add("所有物料");
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                this.comboBox2.Text = "";

                this.comboBox2.Items.Clear();
                this.comboBox3.Items.Clear();
                this.comboBox2.Items.Add("主营");


                this.comboBox3.Items.Add("配件原材料");
                this.comboBox3.Items.Add("结构件（总数）");
                this.comboBox3.Items.Add("包材");
                this.comboBox3.Items.Add("辅料");
                this.comboBox3.Items.Add("开关");
                this.comboBox3.Items.Add("电子结构");
                this.comboBox3.Items.Add("连接器");
                this.comboBox3.Items.Add("线材");
                this.comboBox3.Items.Add("小五金");
                this.comboBox3.Items.Add("橡硅胶");
                this.comboBox3.Items.Add("大五金");
                this.comboBox3.Items.Add("铝壳");
               // this.comboBox3.Items.Add("所有物料");

            }
            else if (comboBox1.SelectedIndex == 2)
            {
                this.comboBox2.Text = "";

                this.comboBox3.Items.Clear();
                this.comboBox2.Items.Clear();
                this.comboBox2.Items.Add("主营");

                this.comboBox3.Items.Add("PCB/FCB");
                this.comboBox3.Items.Add("电容");
                this.comboBox3.Items.Add("电阻");
                this.comboBox3.Items.Add("电感");
                this.comboBox3.Items.Add("二极管");
                this.comboBox3.Items.Add("三极管");
                this.comboBox3.Items.Add("场效应管");
                this.comboBox3.Items.Add("集成电路");
                this.comboBox3.Items.Add("晶振(石英)");
                this.comboBox3.Items.Add("滤波器");
                this.comboBox3.Items.Add("鉴频器/解调器");
                this.comboBox3.Items.Add("保险类");
               // this.comboBox3.Items.Add("所有物料");
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedItem== "所有客户")
            {
                comboBox3.Text = "";
                comboBox3.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = -1;
            //comboBox2.Enabled = true;
            comboBox3.Enabled = true;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox3.SelectedItem== "所有物料")
            {
                comboBox2.Text = "";
                comboBox2.Enabled = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            chartControl1.Series.Clear();
            gridView1.OptionsView.ShowGroupPanel = true;
            gridView1.GroupPanelText ="来料异常分析";
            CreatTitleName();
            DateTime starttime = txtarrivatedate1.Value.Date;
            DateTime endtime = txtarrivatedate2.Value.AddDays(1).Date;
            if (comboBox2.Enabled == true && comboBox3.Enabled==true)
            {
                CreatNormalType(starttime, endtime);
            }
            else if (comboBox2.Enabled==true && comboBox3.Enabled==false)
            {
                CreatCustomerType(starttime, endtime);
            }
        }

        private void CreatTitleName()
        {
            DevExpress.XtraCharts.ChartTitle chartTitle = new DevExpress.XtraCharts.ChartTitle();
            chartTitle.Text = "Pareto Chart of 来料不良类别";//标题内容
            chartTitle.TextColor = Color.Black; //颜色设置
            chartTitle.Font = new Font("Tahoma", 13);//字体类型字号
            chartTitle.Alignment = StringAlignment.Center;
            chartControl1.Titles.Clear();
            chartControl1.Titles.Add(chartTitle);
        }

        private void CreatCustomerType(DateTime starttime, DateTime endtime)
        {
            int n1 = 0;
            int n2 = 0;
            int n3 = 0;
            int n4 = 0;
            int n5 = 0;
            string[] NgType = { "标签", "包装","数量", "错料", "外观", "尺寸", "试装不良", "功能", "环保", "其他" };
            int[] Number = new int[10];
            DataTable dt = new DataTable();
            DataColumn dc1 = new DataColumn("不良类别", Type.GetType("System.String"));
            DataColumn dc2 = new DataColumn("不良批数", Type.GetType("System.Int32"));
            dt.Columns.Add(dc1);
            dt.Columns.Add(dc2);  
            

            DataTable dt1 = new DataTable();
            DataColumn dc3 = new DataColumn("不良类别", Type.GetType("System.String"));
            DataColumn dc6 = new DataColumn("不良批数", Type.GetType("System.Int32"));//0305新添加不良批数。
            DataColumn dc5= new DataColumn("Percent", Type.GetType("System.String"));//0305新添加不良批数对应的百分比。
            DataColumn dc4 = new DataColumn("Cum%", Type.GetType("System.String"));
            dt1.Columns.Add(dc3);
            dt1.Columns.Add(dc6);
            dt1.Columns.Add(dc5);
            dt1.Columns.Add(dc4);
            string constr = "server=192.168.0.176;database=BarcodeNew;user id=sa;password=The0more7people0you7love3the7weaker8you8are;Pooling=false";
            //注意：此处连接了正式生产数据库。
            for (int i = 0; i < 10; i++)
            {
                using (SqlConnection con = new SqlConnection(constr))
                {
                    string sql1 = "QMS_NgTypeforCustomer";
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
                if (Number[i]!=0)
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
                //dr["Pencent"] = n4==0 ? 0.ToString() : (ArrayNumber[i] / (double)n2).ToString("p");
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
            //DIA.Auto = false;
            //DIA.MaxValue = n2;
            //DIA.MinValue = 0;
           // DIA.SetInternalMinMaxValues(0,n2);
            // DIA.SetMinMaxValues(0,n2);minInterval
            // EnableAxisYScrolling = true


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
            //myAxisY.Range.MaxValue = 1.00;
            // myAxisY.VisualRange.MaxValueInternal = 1;
            // myAxisY.VisualRange.MaxValueInternal=1;
            //myAxisY.VisualRange.SetMinMaxValues(0, 1);


            myAxisY.Title.Text = "不良占比";
            myAxisY.Title.Alignment = StringAlignment.Center;
            myAxisY.Title.Visible = true;
            myAxisY.Title.TextColor = Color.Black;
            myAxisY.Title.Font = new Font("Tahoma", 10);
            myAxisY.Title.Antialiasing = true;

            ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
            myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;//显示为百分数
        }

        private void CreatNormalType(DateTime starttime, DateTime endtime)
        {
            int n1 = 0;
            int n2 = 0;
            int n3 = 0;
            string[] NgType = { "包装", "标签", "数量" , "错料", "外观", "尺寸", "试装不良", "功能", "环保", "其他" };
            int[] Number = new int[10];
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
            string constr = "server=192.168.0.176;database=BarcodeNew;user id=sa;password=The0more7people0you7love3the7weaker8you8are;Pooling=false";
            for (int i = 0; i < 10; i++)
            {
                using (SqlConnection con = new SqlConnection(constr))
                {
                    string sql1 = "QMS_NgTypeAnalyse";
                    using (SqlCommand cmd = new SqlCommand(sql1, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@ubussiness", comboBox1.SelectedItem);
                        cmd.Parameters.AddWithValue("@ucustomer", comboBox2.SelectedItem);
                        cmd.Parameters.AddWithValue("@umaterialstype", comboBox3.SelectedItem);
                        cmd.Parameters.AddWithValue("@ustarttime", starttime);
                        cmd.Parameters.AddWithValue("@uendtime", endtime);
                        cmd.Parameters.AddWithValue("@ngtype", NgType[i]);//添加一种不良类别，然后分别统计该类别的不良批数
                        con.Open();
                        Number[i] = (int)cmd.ExecuteScalar();
                    }
                }
                    n2 += Number[i];
                    DataRow dr = dt.NewRow();
                    dr["不良类别"] = NgType[i];
                    dr["不良批数"] = Number[i];
                    //去掉不良类别中不良批数为0的行。
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
                n3 += ArrayNumber[i];
                DataRow dr = dt1.NewRow();
                dr["不良类别"] = dt.Rows[i]["不良类别"];
                dr["不良批数"] = dt.Rows[i]["不良批数"];
                dr["Percent"] = (ArrayNumber[i] / (double)n2).ToString("p");
                //dr["Pencent"] = n4==0 ? 0.ToString() : (ArrayNumber[i] / (double)n2).ToString("p");
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
            //DIA.Auto = false;
            //DIA.SetInternalMinMaxValues(0, n2);
            //DIA.MaxValue = n2;
            //DIA.MinValue = 0;

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

            //myAxisY.AxisRange.Auto
            //myAxisY.Range.Auto = false;
            //myAxisY.Range.MinValue = 0;
            //myAxisY.Range.MaxValue = 1.00;



            myAxisY.Title.Text = "不良占比";
            myAxisY.Title.Alignment = StringAlignment.Center;
            myAxisY.Title.Visible = true;
            myAxisY.Title.TextColor = Color.Black;
            myAxisY.Title.Font = new Font("Tahoma", 10);
            myAxisY.Title.Antialiasing = true;
            //myAxisY.Label.TextColor = Color.Blue;
            //myAxisY.Color = Color.Blue;

            ((LineSeriesView)series3.View).LineMarkerOptions.Color = Color.Red;
            myAxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;//显示为百分数
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
    }
}
