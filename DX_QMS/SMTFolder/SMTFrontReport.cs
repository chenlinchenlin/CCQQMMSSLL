using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraCharts;
using DevExpress.Utils;
using DevExpress.XtraPrinting;
using DevExpress.XtraPrintingLinks;
using DevExpress.XtraBars;
using DX_QMS.Common;
using System.IO;
using DevExpress.XtraGrid.Views.Base;

namespace DX_QMS.SMTFolder
{
    public partial class SMTFrontReport : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public SMTFrontReport()
        {
            InitializeComponent();
        }

        private void SMTFrontReport_Load(object sender, EventArgs e)
        {
            txtstartdate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            txtenddate.Text = DateTime.Now.ToString("yyyy-MM-dd");
            txtreporttype.SelectedIndex = 0;
            selecttype.SelectedIndex = 0;
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }

        private void txtstartdate_EditValueChanged(object sender, EventArgs e)
        {
            selecttype_SelectedIndexChanged(sender,e);
        }

        private void selecttype_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (selecttype.SelectedIndex == 0)
            {
                txtreporttype.SelectedIndex = 0;
                //txtreporttype.Enabled = false;
                txtmodel.Text = "";
                txtmodel.Enabled = false;
                txtworkno.Enabled = true;
                txtworkno.Focus();
                lblworkno.Text = "多个工单号用分号或者逗号分割";
                lblmodel.Text = "";
                checktype.DataSource = null;
                checktype.Enabled = false;

            }
            else if (selecttype.SelectedIndex == 1)
            {
                txtreporttype.SelectedIndex = 0;
                txtreporttype.Enabled = true;
                txtworkno.Text = "";
                txtworkno.Enabled = false;
                txtmodel.Enabled = true;
                txtmodel.Focus();
                lblworkno.Text = "";
                lblmodel.Text = "多个机型用分号或者逗号分割";
                checktype.DataSource = null;
                checktype.Enabled = false;

            }
            if (selecttype.SelectedIndex == 2)
            {
                txtreporttype.SelectedIndex = 0;
                txtreporttype.Enabled = true;
                txtworkno.Text = "";
                lblworkno.Text = "";
                lblmodel.Text = "";
                txtworkno.Enabled = false;
                txtmodel.Text = "";
                txtmodel.Enabled = false;
                checktype.Enabled = true;
                DataTable dt = null;
                string sql = "  select distinct customer from OEM_Prod where eventtime >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and eventtime <=  '" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  and eventtime is not null and customer<>''  ";
                dt = DbAccess.SelectBySql(sql).Tables[0];
                checktype.DataSource = dt;
                checktype.DisplayMember = dt.Columns["customer"].ToString();
                checktype.ValueMember = dt.Columns["customer"].ToString();
            }
            else if (selecttype.SelectedIndex == 3)
            {
                txtreporttype.SelectedIndex = 0;
                txtreporttype.Enabled = true;
                txtworkno.Text = "";
                txtworkno.Enabled = false;
                lblworkno.Text = "";
                lblmodel.Text = "";
                txtmodel.Text = "";
                txtmodel.Enabled = false;
                checktype.Enabled = true;
                DataTable dt = null;
                string sql = " select distinct l.laName 线别 from SMT_Prod s left join latype l on s.latype = l.la  where eventtime >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and eventtime <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59' and eventtime is not null  ";
                dt = DbAccess.SelectBySql(sql).Tables[0];
                checktype.DataSource = dt;
                checktype.DisplayMember = dt.Columns["线别"].ToString();
                checktype.ValueMember = dt.Columns["线别"].ToString();
            }

        }

        private void sBtnselect_Click(object sender, EventArgs e)
        {
            if (txtstartdate.Text == "" || txtenddate.Text == "")
                return;


            if (selecttype.SelectedIndex == 0)  //工单
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
                    workno += " or  workno ='" + sArray[i]+ "'";
                }
                if (txtreporttype.Text == "日报")
                {
                    dt = worknodayReport(workno);
                }
                else if (txtreporttype.Text == "周报")
                {
                    dt = worknoweekReport(workno);
                }
                else if (txtreporttype.Text == "月报")
                {
                    dt = worknomonthReport(workno);
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
                    Series series = CreateWorknoSeries(sArray[i], ViewType.Line, dt.Select("工单号='" + sArray[i] + "'"));
                    chartControl.Series.Add(series);
                }

                ((XYDiagram)chartControl.Diagram).AxisY.Title.Visibility = DefaultBoolean.True;
                ((XYDiagram)chartControl.Diagram).AxisY.Title.Text = "良率";
                ((XYDiagram)chartControl.Diagram).AxisY.Color = Color.Gray;
                ((XYDiagram)chartControl.Diagram).AxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;
                AxisRange DIA = (AxisRange)((XYDiagram)chartControl.Diagram).AxisY.Range;
                DIA.SetMinMaxValues(0.95, 1);

            }
            else if (selecttype.SelectedIndex == 1)  // 机型
            {
                if (txtmodel.Text.Trim() == "")
                    return;
                DataTable dt = new DataTable ();
                DataTable modeldt = null;
                string str = txtmodel.Text.Trim();
                string[] sArray = str.Split(new char[4] { ';', '；', ',', '，' });
                ChartTitle chartTitle = new ChartTitle();
                chartTitle.Text = txtstartdate.DateTime.ToString("yyyy-MM-dd") + "至" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " " +str+ "机型的" + txtreporttype.Text + "良率情况";
                chartTitle.TextColor = System.Drawing.Color.Black;
                chartTitle.Font = new Font("微软雅黑", 10);
                chartTitle.Alignment = StringAlignment.Center;
                chartControl.Titles.Clear();
                chartControl.Titles.Add(chartTitle);
                chartControl.Series.Clear();

               if (txtreporttype.Text == "日报")
               {
                    for (int i = 0; i < sArray.Length; i++)
                    {
                        if (sArray[i].Trim() == "")
                            continue;
                        modeldt = null;                       
                        modeldt = modeldayReport(sArray[i]);
                        Series series = CreatemodelSeries(sArray[i], ViewType.Line, modeldt);
                        chartControl.Series.Add(series);
                        dt.Merge(modeldt);
                    }
                }
                else if (txtreporttype.Text == "周报")
                {
                    for (int i = 0; i < sArray.Length; i++)
                    {
                        if (sArray[i].Trim() == "")
                            continue;
                        modeldt = null;                     
                        modeldt = modelweekReport(sArray[i]);
                        Series series = CreatemodelSeries(sArray[i], ViewType.Line, modeldt);
                        chartControl.Series.Add(series);
                        dt.Merge(modeldt);
                    }
                }
                else if (txtreporttype.Text == "月报")
                {
                    for (int i = 0; i < sArray.Length; i++)
                    {
                        if (sArray[i].Trim() == "")
                            continue;
                        modeldt = null;                  
                        modeldt = modelmonthReport(sArray[i]);
                        Series series = CreatemodelSeries(sArray[i], ViewType.Line, modeldt);
                        chartControl.Series.Add(series);
                        dt.Merge(modeldt);
                    }
                }

                ((XYDiagram)chartControl.Diagram).AxisY.Title.Visibility = DefaultBoolean.True;
                ((XYDiagram)chartControl.Diagram).AxisY.Title.Text = "良率";
                ((XYDiagram)chartControl.Diagram).AxisY.Color = Color.Gray;
                ((XYDiagram)chartControl.Diagram).AxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;
                AxisRange DIA = (AxisRange)((XYDiagram)chartControl.Diagram).AxisY.Range;
                DIA.SetMinMaxValues(0.96, 1);

                gridControl.DataSource = null;
                gridView.Columns.Clear();
                gridControl.DataSource = dt;
            }
            else if (selecttype.SelectedIndex == 2) //客户   
            {

                DataTable dt = null;
                string customer = "";
                if (checktype.CheckedItems.Count < 1)
                    return;

                for (int i = 0; i < checktype.CheckedItems.Count; i++)
                {
                    customer += ",'" + checktype.CheckedItems[i].ToString() + "'";
                }
                customer = customer.TrimStart(',');

                if (txtreporttype.Text == "日报")
                {
                    dt = customerdayReport(customer);
                }
                else if (txtreporttype.Text == "周报")
                {
                    dt = customerweekReport(customer, "周数");
                }
                else if (txtreporttype.Text == "月报")
                {
                    dt = customermonthReport(customer, "月份");
                }
                gridControl.DataSource = null;
                gridView.Columns.Clear();
                gridControl.DataSource = dt;

                ChartTitle chartTitle = new ChartTitle();
                chartTitle.Text = txtstartdate.DateTime.ToString("yyyy-MM-dd") + "至" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " " + customer + "客户的" + txtreporttype.Text + "良率情况";
                chartTitle.TextColor = System.Drawing.Color.Black; //颜色设置
                chartTitle.Font = new Font("微软雅黑", 10);
                chartTitle.Alignment = StringAlignment.Center;
                chartControl.Titles.Clear();
                chartControl.Titles.Add(chartTitle);
                chartControl.Series.Clear();

                for (int i = 0; i < checktype.CheckedItems.Count; i++)
                {
                    Series series = CreatecustomerSeries(checktype.CheckedItems[i].ToString(), ViewType.Line, dt.Select("客户='" + checktype.CheckedItems[i].ToString() + "'"));
                    chartControl.Series.Add(series);
                }

                ((XYDiagram)chartControl.Diagram).AxisY.Title.Visibility = DefaultBoolean.True;
                ((XYDiagram)chartControl.Diagram).AxisY.Title.Text = "良率";
                ((XYDiagram)chartControl.Diagram).AxisY.Color = Color.Gray;
                ((XYDiagram)chartControl.Diagram).AxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;
                AxisRange DIA = (AxisRange)((XYDiagram)chartControl.Diagram).AxisY.Range;
                DIA.SetMinMaxValues(0.96, 1);
            }
            else if (selecttype.SelectedIndex == 3)  // 线体
            {
                DataTable dt = null;
                string latype = "";
                if (checktype.CheckedItems.Count < 1)
                    return;

                for (int i = 0; i < checktype.CheckedItems.Count; i++)
                {
                    latype += ",'" + checktype.CheckedItems[i].ToString() + "'";
                }
                latype = latype.TrimStart(',');

                if (txtreporttype.Text == "日报")
                {
                    dt = latypedayReport(latype);            
                }
                else if (txtreporttype.Text == "周报")
                {
                    dt = latypeweekReport(latype);                 
                }
                else if (txtreporttype.Text == "月报")
                {
                    dt = latypemonthReport(latype);
                }
                gridControl.DataSource = null;
                gridView.Columns.Clear();             
                gridControl.DataSource = dt;

                ChartTitle chartTitle = new ChartTitle();
                chartTitle.Text = txtstartdate.DateTime.ToString("yyyy-MM-dd") + "至" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " " + latype + "线别的"+txtreporttype.Text+ "良率情况";
                chartTitle.TextColor = System.Drawing.Color.Black; //颜色设置
                chartTitle.Font = new Font("微软雅黑", 10);
                chartTitle.Alignment = StringAlignment.Center;
                chartControl.Titles.Clear();
                chartControl.Titles.Add(chartTitle);
                chartControl.Series.Clear();

                for (int i = 0; i < checktype.CheckedItems.Count; i++)
                {
                    Series series = CreateLatypeSeries(checktype.CheckedItems[i].ToString(), ViewType.Line, dt.Select("线别='" +checktype.CheckedItems[i].ToString()+ "'"));
                    chartControl.Series.Add(series);
                }

                ((XYDiagram)chartControl.Diagram).AxisY.Title.Visibility = DefaultBoolean.True;
                ((XYDiagram)chartControl.Diagram).AxisY.Title.Text = "良率";
                ((XYDiagram)chartControl.Diagram).AxisY.Color = Color.Gray;
                ((XYDiagram)chartControl.Diagram).AxisY.NumericOptions.Format = DevExpress.XtraCharts.NumericFormat.Percent;
                AxisRange DIA = (AxisRange)((XYDiagram)chartControl.Diagram).AxisY.Range;
                DIA.SetMinMaxValues(0.8, 1);

            }

         }

        DataTable worknodayReport(string workno)
        {
            DataTable dt = null;
            string sql = "   select  z.org_id 组织,z.workno 工单号,isnull(sum(s.NGqty),0) NG数量,sum(z.qty) 送板数量,convert(numeric(5,4),((sum(z.qty)-isnull(sum(s.NGqty),0)+0.0)/(sum(z.qty)))) as 合格率,z.operDate 日期     ";
            sql += "  from (  select  org_id,workno,sum(NGqty) NGqty,CONVERT(varchar(100),cdate, 23) operDate from   ";
            sql += " ( select 'SHL' org_id,workNo workno,1 NGqty,illDate cdate from OEM_Repair where  type='炉后和AOI检查'  and illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'   ";
            sql += "  union all  select a.org_id,a.workno ,a.qty NGqty,min(a.checkDate) cdate  from SMT_WorknoRepair a  where a.station in('AOI检测','炉后QC检测') and a.checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and a.checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'   group by a.org_id,a.workno,a.qty   ";
            sql += "  ) n group by org_id,workno,CONVERT(varchar(100),cdate, 23)  ) s   right join    (    ";
            sql += "  select  org_id,workno,sum(qty) qty ,CONVERT(varchar(100),operDate, 23) operDate from SMT_SONGBAN where ("+ workno+ ") and qty>0 and operDate >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and operDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  ";
            sql += "  group by org_id,workno, CONVERT(varchar(100),operDate, 23)  ) z  on s.org_id = z.org_id and s.workno = z.workno and s.operDate = z.operDate  group by  z.org_id,z.workno,z.operDate   ";     
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable worknoweekReport(string workno)
        {
            DataTable dt = null;
            string sql = "   select  z.org_id 组织,z.workno 工单号,isnull(sum(s.NGqty),0) NG数量,sum(z.qty) 送板数量,convert(numeric(5,4),((sum(z.qty)-isnull(sum(s.NGqty),0)+0.0)/(sum(z.qty)))) as 合格率,z.operDate 周数     ";
            sql += "  from (  select  org_id,workno,sum(NGqty) NGqty,dbo.getWeekNoBtDate(cdate)  operDate from   ";
            sql += " ( select 'SHL' org_id,workNo workno,1 NGqty,illDate cdate from OEM_Repair where  type='炉后和AOI检查'  and illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'   ";
            sql += "  union all  select a.org_id,a.workno ,a.qty NGqty,min(a.checkDate) cdate  from SMT_WorknoRepair a  where a.station in('AOI检测','炉后QC检测') and a.checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and a.checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'   group by a.org_id,a.workno,a.qty   ";
            sql += "  ) n group by org_id,workno,dbo.getWeekNoBtDate(cdate)   ) s   right join    (    ";
            sql += "  select  org_id,workno,sum(qty) qty ,dbo.getWeekNoBtDate(operDate) operDate from SMT_SONGBAN where (" + workno + ") and qty>0 and operDate >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and operDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  ";
            sql += "  group by org_id,workno, dbo.getWeekNoBtDate(operDate)  ) z  on s.org_id = z.org_id and s.workno = z.workno and s.operDate = z.operDate  group by  z.org_id,z.workno,z.operDate   ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable worknomonthReport(string workno)
        {
            DataTable dt = null;
            string sql = "   select  z.org_id 组织,z.workno 工单号,isnull(sum(s.NGqty),0) NG数量,sum(z.qty) 送板数量,convert(numeric(5,4),((sum(z.qty)-isnull(sum(s.NGqty),0)+0.0)/(sum(z.qty)))) as 合格率,z.operDate 月份     ";
            sql += "  from (  select  org_id,workno,sum(NGqty) NGqty,dbo.getMonthByDate(cdate)  operDate from   ";
            sql += " ( select 'SHL' org_id,workNo workno,1 NGqty,illDate cdate from OEM_Repair where  type='炉后和AOI检查'  and illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'   ";
            sql += "  union all  select a.org_id,a.workno ,a.qty NGqty,min(a.checkDate) cdate  from SMT_WorknoRepair a  where a.station in('AOI检测','炉后QC检测') and a.checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and a.checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'   group by a.org_id,a.workno,a.qty   ";
            sql += "  ) n group by org_id,workno,dbo.getMonthByDate(cdate)   ) s   right join    (    ";
            sql += "  select  org_id,workno,sum(qty) qty ,dbo.getMonthByDate(operDate) operDate from SMT_SONGBAN where (" + workno + ") and qty>0 and operDate >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and operDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  ";
            sql += "  group by org_id,workno, dbo.getMonthByDate(operDate)  ) z  on s.org_id = z.org_id and s.workno = z.workno and s.operDate = z.operDate  group by  z.org_id,z.workno,z.operDate   ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable latypedayReport(string latype)
        {
            DataTable dt = null;
            string sql = "  select  isnull(sum(s.NGqty),0) NG数量,sum(z.qty) 数量,convert(numeric(5,4),((sum(z.qty)-isnull(sum(s.NGqty),0)+0.0)/(sum(z.qty)))) as 合格率,z.laName 线别,z.operDate 日期   ";
            sql += " from (  select  org_id,workno,sum(NGqty) NGqty,CONVERT(varchar(100),cdate, 23) operDate from  (  ";
            sql += " select 'SHL' org_id,workNo workno,1 NGqty,illDate cdate from OEM_Repair where  type='炉后和AOI检查'  and illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59' ";
            sql += "   union all  select a.org_id,a.workno ,a.qty NGqty,min(a.checkDate) cdate  from SMT_WorknoRepair a  where a.station in('AOI检测','炉后QC检测') and a.checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and a.checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  ";
            sql += " group by a.org_id,a.workno,a.qty  ) n group by org_id,workno,CONVERT(varchar(100),cdate, 23)	) s  right join  (   ";
            sql += "  select m.org_id ,m.workno,sum(m.qty) qty,w.laName,CONVERT(varchar(100),m.operDate, 23) operDate from SMT_SONGBAN  m  left join ( select org_id,workno,l.laName from  ";
            sql += "  (  select org_id, workno,latype from   ( select  org_id,workno,latype from OEM_Prod union all  select  org_id,workno,latype from SMT_Prod   ) w  where w.latype <>'' and w.latype is not null  group by org_id,workno,latype  ";
            sql += "  ) s left join latype l on s.latype = l.la  where l.laName in (" + latype + ") ) w    on  m.org_id = w.org_id and m.workno = w.workno  where m.qty >0  and  m.operDate >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and m.operDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  and w.laName is not null  ";
            sql += "  group by m.org_id ,m.workno,w.laName,CONVERT(varchar(100),m.operDate, 23)   ) z  on s.org_id = z.org_id and s.workno = z.workno and s.operDate = z.operDate  group by  z.laName,z.operDate  ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable latypeweekReport(string latype)
        {
            DataTable dt = null;
            string sql = "  select  isnull(sum(s.NGqty),0) NG数量,sum(z.qty) 数量,convert(numeric(5,4),((sum(z.qty)-isnull(sum(s.NGqty),0)+0.0)/(sum(z.qty)))) as 合格率,z.laName 线别,z.operDate 周数   ";
            sql += " from (  select  org_id,workno,sum(NGqty) NGqty,dbo.getWeekNoBtDate(cdate) operDate from  (  ";
            sql += " select 'SHL' org_id,workNo workno,1 NGqty,illDate cdate from OEM_Repair where  type='炉后和AOI检查'  and illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59' ";
            sql += "   union all  select a.org_id,a.workno ,a.qty NGqty,min(a.checkDate) cdate  from SMT_WorknoRepair a  where a.station in('AOI检测','炉后QC检测') and a.checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and a.checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  ";
            sql += " group by a.org_id,a.workno,a.qty  ) n group by org_id,workno,dbo.getWeekNoBtDate(cdate)	) s  right join  (   ";
            sql += "  select m.org_id ,m.workno,sum(m.qty) qty,w.laName,dbo.getWeekNoBtDate(m.operDate) operDate from SMT_SONGBAN  m  left join ( select org_id,workno,l.laName from  ";
            sql += "  (  select org_id, workno,latype from   ( select  org_id,workno,latype from OEM_Prod union all  select  org_id,workno,latype from SMT_Prod   ) w  where w.latype <>'' and w.latype is not null  group by org_id,workno,latype  ";
            sql += "  ) s left join latype l on s.latype = l.la  where l.laName in (" + latype + ") ) w    on  m.org_id = w.org_id and m.workno = w.workno  where m.qty >0  and  m.operDate >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and m.operDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  and w.laName is not null  ";
            sql += "  group by m.org_id ,m.workno,w.laName,dbo.getWeekNoBtDate(m.operDate)   ) z  on s.org_id = z.org_id and s.workno = z.workno and s.operDate = z.operDate  group by  z.laName,z.operDate  ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable latypemonthReport(string latype)
        {
            DataTable dt = null;
            string sql = "  select  isnull(sum(s.NGqty),0) NG数量,sum(z.qty) 数量,convert(numeric(5,4),((sum(z.qty)-isnull(sum(s.NGqty),0)+0.0)/(sum(z.qty)))) as 合格率,z.laName 线别,z.operDate 月份   ";
            sql += " from (  select  org_id,workno,sum(NGqty) NGqty,dbo.getMonthByDate(cdate) operDate from  (  ";
            sql += " select 'SHL' org_id,workNo workno,1 NGqty,illDate cdate from OEM_Repair where  type='炉后和AOI检查'  and illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59' ";
            sql += "   union all  select a.org_id,a.workno ,a.qty NGqty,min(a.checkDate) cdate  from SMT_WorknoRepair a  where a.station in('AOI检测','炉后QC检测') and a.checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and a.checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  ";
            sql += " group by a.org_id,a.workno,a.qty  ) n group by org_id,workno,dbo.getMonthByDate(cdate)	) s  right join  (   ";
            sql += "  select m.org_id ,m.workno,sum(m.qty) qty,w.laName,dbo.getMonthByDate(m.operDate) operDate from SMT_SONGBAN  m  left join ( select org_id,workno,l.laName from  ";
            sql += "  (  select org_id, workno,latype from   ( select  org_id,workno,latype from OEM_Prod union all  select  org_id,workno,latype from SMT_Prod   ) w  where w.latype <>'' and w.latype is not null  group by org_id,workno,latype  ";
            sql += "  ) s left join latype l on s.latype = l.la  where l.laName in ("+latype+ ") ) w    on  m.org_id = w.org_id and m.workno = w.workno  where m.qty >0  and  m.operDate >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and m.operDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  and w.laName is not null  ";
            sql += "  group by m.org_id ,m.workno,w.laName,dbo.getMonthByDate(m.operDate)   ) z  on s.org_id = z.org_id and s.workno = z.workno and s.operDate = z.operDate  group by  z.laName,z.operDate  ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable customerdayReport(string customer)
        {
            DataTable dt = null;
            string sql = " 		select  isnull(sum(s.NGqty),0) NG数量,sum(z.qty) 送板数量,convert(numeric(5,4),((sum(z.qty)-isnull(sum(s.NGqty),0)+0.0)/(sum(z.qty)))) as 合格率,z.customer 客户,z.operDate 日期    ";
            sql += "  from (  select org_id,workno,sum(NGqty) NGqty,CONVERT(varchar(100),cdate, 23) operDate  from  (   ";
            sql += "  select 'SHL' org_id,workNo workno,1 NGqty,illDate cdate from OEM_Repair where  type='炉后和AOI检查'  and illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'     ";
            sql += " union all select a.org_id,a.workno,a.qty NGqty,min(a.checkDate) cdate  from SMT_WorknoRepair a  where a.station in('AOI检测','炉后QC检测') and a.checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and a.checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'   group by a.org_id,a.workno,a.qty   ";
            sql += "  ) n group by org_id,workno,CONVERT(varchar(100),cdate, 23)  	) s   right join  (      ";
            sql += " select  s.org_id,s.workno,g.customer customer,sum(s.qty) qty, CONVERT(varchar(100),s.operDate, 23) operDate from ( select  org_id,workno,customer from OEM_Prod  where customer in (" + customer + ") group by org_id,workno,customer ) g    ";
            sql += " left join SMT_SONGBAN s on g.org_id = s.org_id and g.workno = s.workno where s.qty is not null and s.qty >0  and  operDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and operDate <= '" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  ";
            sql += " group by  s.org_id,s.workno,g.customer, CONVERT(varchar(100),s.operDate, 23)   ";
            sql += " ) z  on s.org_id = z.org_id and s.workno = z.workno  and s.operDate = z.operDate    ";
            sql += " 	group by  z.customer,z.operDate  order by z.operDate asc   ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable customerweekReport(string customer, string time)
        {
            DataTable dt = null;
            string sql = " 		select  isnull(sum(s.NGqty),0) NG数量,sum(z.qty) 送板数量,convert(numeric(5,4),((sum(z.qty)-isnull(sum(s.NGqty),0)+0.0)/(sum(z.qty)))) as 合格率,z.customer 客户,dbo.getWeekNoBtDate(z.operDate) 周数    ";
            sql += "  from (  select org_id,workno,sum(NGqty) NGqty,CONVERT(varchar(100),cdate, 23) operDate  from  (   ";
            sql += "  select 'SHL' org_id,workNo workno,1 NGqty,illDate cdate from OEM_Repair where  type='炉后和AOI检查'  and illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'     ";
            sql += " union all select a.org_id,a.workno,a.qty NGqty,min(a.checkDate) cdate  from SMT_WorknoRepair a  where a.station in('AOI检测','炉后QC检测') and a.checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and a.checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'   group by a.org_id,a.workno,a.qty   ";
            sql += "  ) n group by org_id,workno,CONVERT(varchar(100),cdate, 23)  	) s   right join  (      ";
            sql += " select  s.org_id,s.workno,g.customer customer,sum(s.qty) qty, CONVERT(varchar(100),s.operDate, 23) operDate from ( select  org_id,workno,customer from OEM_Prod  where customer in (" + customer + ") group by org_id,workno,customer ) g    ";
            sql += " left join SMT_SONGBAN s on g.org_id = s.org_id and g.workno = s.workno where s.qty is not null and s.qty >0  and  operDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and operDate <= '" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  ";
            sql += " group by  s.org_id,s.workno,g.customer, CONVERT(varchar(100),s.operDate, 23)   ";
            sql += " ) z  on s.org_id = z.org_id and s.workno = z.workno  and s.operDate = z.operDate    ";
            sql += " 	group by  z.customer,dbo.getWeekNoBtDate(z.operDate)  order by dbo.getWeekNoBtDate(z.operDate) asc   ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable customermonthReport(string customer, string time)
        {
            DataTable dt = null;
            string sql = " 		select  isnull(sum(s.NGqty),0) NG数量,sum(z.qty) 送板数量,convert(numeric(5,4),((sum(z.qty)-isnull(sum(s.NGqty),0)+0.0)/(sum(z.qty)))) as 合格率,z.customer 客户,dbo.getMonthByDate(z.operDate) 月份    ";
            sql += "  from (  select org_id,workno,sum(NGqty) NGqty,CONVERT(varchar(100),cdate, 23) operDate  from  (   ";
            sql += "  select 'SHL' org_id,workNo workno,1 NGqty,illDate cdate from OEM_Repair where  type='炉后和AOI检查'  and illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'     ";
            sql += " union all select a.org_id,a.workno,a.qty NGqty,min(a.checkDate) cdate  from SMT_WorknoRepair a  where a.station in('AOI检测','炉后QC检测') and a.checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and a.checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'   group by a.org_id,a.workno,a.qty   ";
            sql += "  ) n group by org_id,workno,CONVERT(varchar(100),cdate, 23)  	) s   right join  (      ";
            sql += " select  s.org_id,s.workno,g.customer customer,sum(s.qty) qty, CONVERT(varchar(100),s.operDate, 23) operDate from ( select  org_id,workno,customer from OEM_Prod  where customer in (" + customer + ") group by org_id,workno,customer ) g    ";
            sql += " left join SMT_SONGBAN s on g.org_id = s.org_id and g.workno = s.workno where s.qty is not null and s.qty >0  and  operDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and operDate <= '" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  ";
            sql += " group by  s.org_id,s.workno,g.customer, CONVERT(varchar(100),s.operDate, 23)   ";
            sql += " ) z  on s.org_id = z.org_id and s.workno = z.workno  and s.operDate = z.operDate    ";
            sql += " 	group by  z.customer,dbo.getMonthByDate(z.operDate)  order by dbo.getMonthByDate(z.operDate) asc   ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }


        DataTable modeldayReport(string model)
        {
            DataTable dt = null;
            string sql = " select  isnull(sum(s.NGqty),0) NG数量,sum(z.qty) 送板数量,convert(numeric(5,4),((sum(z.qty)-isnull(sum(s.NGqty),0)+0.0)/(sum(z.qty)))) as 合格率,'"+ model+ "' 编码,z.operDate 日期     ";
            sql += " from (  select  org_id, workno,sum(NGqty) NGqty,CONVERT(varchar(100),cdate, 23) operDate from    ";
            sql += " ( select 'SHL' org_id,workNo workno,1 NGqty,illDate cdate from OEM_Repair where  type='炉后和AOI检查'  and illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'   ";
            sql += " union all select a.org_id,a.workno ,a.qty NGqty,min(a.checkDate) cdate  from SMT_WorknoRepair a  where a.station in('AOI检测','炉后QC检测') and a.checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and a.checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59' ";
            sql += " group by a.org_id,a.workno,a.qty  ) n group by org_id,workno,CONVERT(varchar(100),cdate, 23)  ) s  right join  (   ";
            sql += " select g.org_id org_id,g.workno workno,sum(s.qty) qty, CONVERT(varchar(100),s.operDate, 23) operDate from  (  select org_id, workno,productcode from   ";
            sql += " ( select  org_id,workno,productcode from OEM_Prod  where productcode = '"+ model + "'  union all    select  org_id,workno,productcode from SMT_Prod  where  productcode = '"+ model+ "') w  group by org_id,workno,productcode    ";
            sql += "  ) g  left join SMT_SONGBAN s on g.org_id = s.org_id and g.workno = s.workno where s.qty is not null and s.qty >0 and  operDate >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and operDate <= '" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  group by  g.org_id,g.workno, CONVERT(varchar(100),s.operDate, 23)  ";
            sql += " ) z   on s.org_id = z.org_id and s.workno = z.workno	 and s.operDate = z.operDate  group by  z.operDate    ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable modelweekReport(string model )
        {
            DataTable dt = null;
            string sql = " select  isnull(sum(s.NGqty),0) NG数量,sum(z.qty) 送板数量,convert(numeric(5,4),((sum(z.qty)-isnull(sum(s.NGqty),0)+0.0)/(sum(z.qty)))) as 合格率,'" + model + "' 编码,z.operDate 周数     ";
            sql += " from (  select  org_id, workno,sum(NGqty) NGqty,dbo.getWeekNoBtDate(cdate) operDate from    ";
            sql += " ( select 'SHL' org_id,workNo workno,1 NGqty,illDate cdate from OEM_Repair where  type='炉后和AOI检查'  and illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'   ";
            sql += " union all select a.org_id,a.workno ,a.qty NGqty,min(a.checkDate) cdate  from SMT_WorknoRepair a  where a.station in('AOI检测','炉后QC检测') and a.checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and a.checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59' ";
            sql += " group by a.org_id,a.workno,a.qty  ) n group by org_id,workno,dbo.getWeekNoBtDate(cdate)  ) s  right join  (   ";
            sql += " select g.org_id org_id,g.workno workno,sum(s.qty) qty, dbo.getWeekNoBtDate(s.operDate) operDate from  (  select org_id, workno,productcode from   ";
            sql += " ( select  org_id,workno,productcode from OEM_Prod  where productcode = '" + model + "'  union all    select  org_id,workno,productcode from SMT_Prod  where  productcode = '" + model + "') w  group by org_id,workno,productcode    ";
            sql += "  ) g  left join SMT_SONGBAN s on g.org_id = s.org_id and g.workno = s.workno where s.qty is not null and s.qty >0 and  operDate >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and operDate <= '" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  group by  g.org_id,g.workno,dbo.getWeekNoBtDate(s.operDate)  ";
            sql += " ) z   on s.org_id = z.org_id and s.workno = z.workno	 and s.operDate = z.operDate  group by  z.operDate    ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable modelmonthReport(string model)
        {
            DataTable dt = null;
            string sql = " select  isnull(sum(s.NGqty),0) NG数量,sum(z.qty) 送板数量,convert(numeric(5,4),((sum(z.qty)-isnull(sum(s.NGqty),0)+0.0)/(sum(z.qty)))) as 合格率,'" + model + "' 编码,z.operDate 周数     ";
            sql += " from (  select  org_id, workno,sum(NGqty) NGqty,dbo.getMonthByDate(cdate) operDate from    ";
            sql += " ( select 'SHL' org_id,workNo workno,1 NGqty,illDate cdate from OEM_Repair where  type='炉后和AOI检查'  and illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'   ";
            sql += " union all select a.org_id,a.workno ,a.qty NGqty,min(a.checkDate) cdate  from SMT_WorknoRepair a  where a.station in('AOI检测','炉后QC检测') and a.checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and a.checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59' ";
            sql += " group by a.org_id,a.workno,a.qty  ) n group by org_id,workno,dbo.getMonthByDate(cdate)  ) s  right join  (   ";
            sql += " select g.org_id org_id,g.workno workno,sum(s.qty) qty, dbo.getMonthByDate(s.operDate) operDate from  (  select org_id, workno,productcode from   ";
            sql += " ( select  org_id,workno,productcode from OEM_Prod  where productcode = '" + model + "'  union all    select  org_id,workno,productcode from SMT_Prod  where  productcode = '" + model + "') w  group by org_id,workno,productcode    ";
            sql += "  ) g  left join SMT_SONGBAN s on g.org_id = s.org_id and g.workno = s.workno where s.qty is not null and s.qty >0 and  operDate >= '" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and operDate <= '" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  group by  g.org_id,g.workno,dbo.getMonthByDate(s.operDate)  ";
            sql += " ) z   on s.org_id = z.org_id and s.workno = z.workno	 and s.operDate = z.operDate  group by  z.operDate    ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        private Series CreateWorknoSeries(string caption, ViewType viewType, DataRow[] rw)
        {
            Series series = new Series(caption, viewType);
            foreach (DataRow dr in rw)
            {
                string argument = Convert.ToString(dr[5]);
                double value = Convert.ToDouble(dr[4]);
                series.Points.Add(new SeriesPoint(argument, value));
            }
            series.PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;
            series.ArgumentScaleType = ScaleType.Qualitative;
            series.LabelsVisibility = DefaultBoolean.True;
            return series;
        }

        private Series CreateLatypeSeries(string caption, ViewType viewType, DataRow[] rw)
        {
            Series series = new Series(caption, viewType);

            foreach (DataRow dr in rw)
            {
                string argument = Convert.ToString(dr[4]);
                double value = Convert.ToDouble(dr[2]);
                series.Points.Add(new SeriesPoint(argument, value));
            }
            series.PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;
            series.ArgumentScaleType = ScaleType.Qualitative;
            series.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
            return series;
        }


        private Series CreatecustomerSeries(string caption, ViewType viewType, DataRow[] rw)
        {
            Series series = new Series(caption, viewType);

            foreach (DataRow dr in rw)
            {
                string argument = Convert.ToString(dr[4]);
                double value = Convert.ToDouble(dr[2]);
                series.Points.Add(new SeriesPoint(argument, value));
            }
            series.PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;
            series.ArgumentScaleType = ScaleType.Qualitative;
            series.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
            return series;
        }

        private Series CreatemodelSeries(string caption, ViewType viewType,DataTable dt)
        {
            Series series = new Series(caption, viewType);

            foreach (DataRow dr in dt.Rows)
            {
                string argument = Convert.ToString(dr[4]);
                double value = Convert.ToDouble(dr[2]);
                series.Points.Add(new SeriesPoint(argument, value));
            }
            series.PointOptions.ValueNumericOptions.Format = NumericFormat.Percent;
            series.ArgumentScaleType = ScaleType.Qualitative;
            series.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
            return series;
        }


        private void sBtnreset_Click(object sender, EventArgs e)
        {
            checktype.DataSource = null;
            gridControl.DataSource = null;
            gridView.Columns.Clear();
            chartControl.Titles.Clear();
            chartControl.Series.Clear();
            chartControl.DataSource = null;
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
            string filename = chartControl.Titles[0].ToString();
            ExportToExcel(filename, gridControl, chartControl);
        }

        private void BtnBadClassDetails_Click(object sender, EventArgs e)
        {
            if (txtstartdate.Text == "" || txtenddate.Text == "")
                return;      
              if (selecttype.SelectedIndex == 0)  //工单
              {
                  if (txtworkno.Text.Trim() == "")
                      return;
                  DataTable dt = null;
                  string workno = " where  1=2 ";
                  string str = txtworkno.Text.Trim();
                  string[] sArray = str.Split(new char[4] { ';', '；', ',', '，' });
                  for (int i = 0; i < sArray.Length; i++)
                  {
                      if (sArray[i].Trim() == "")
                          continue;
                      workno += " or  workno ='" + sArray[i] + "'";
                  }
                  dt = worknoBadclassDetail(workno);
                  gridControl.DataSource = null;
                  gridView.Columns.Clear();
                  gridControl.DataSource = dt;

              }
              else if (selecttype.SelectedIndex == 1) // 机型
              {

                  if (txtmodel.Text.Trim() == "")
                      return;
                  DataTable dt = new DataTable();
                  DataTable modeldt = null;
                  string str = txtmodel.Text.Trim();
                  string[] sArray = str.Split(new char[4] { ';', '；', ',', '，' });                     
                      for (int i = 0; i < sArray.Length; i++)
                      {
                          if (sArray[i].Trim() == "")
                              continue;
                          modeldt = null;
                          modeldt = modelBadclassDetail(sArray[i]);
                          dt.Merge(modeldt);
                      }               
                  gridControl.DataSource = null;
                  gridView.Columns.Clear();
                  gridControl.DataSource = dt;
              }
              else if (selecttype.SelectedIndex == 2) // 客户
              {
                  DataTable dt = null;
                  string customer = "";
                  if (checktype.CheckedItems.Count < 1)
                      return;

                  for (int i = 0; i < checktype.CheckedItems.Count; i++)
                  {
                      customer += ",'" + checktype.CheckedItems[i].ToString() + "'";
                  }
                  customer = customer.TrimStart(',');
                  dt = customerBadclassDetail(customer);

                  gridControl.DataSource = null;
                  gridView.Columns.Clear();
                  gridControl.DataSource = dt;

              }
              else if (selecttype.SelectedIndex == 3) // 线别
              {
                  DataTable dt = null;
                string latype = "";
                if (checktype.CheckedItems.Count < 1)
                    return;

                for (int i = 0; i < checktype.CheckedItems.Count; i++)
                {
                    latype += ",'" + checktype.CheckedItems[i].ToString() + "'";
                }
                latype = latype.TrimStart(',');
                dt = latypeBadclassDetail(latype);
                gridControl.DataSource = null;
                gridView.Columns.Clear();
                gridControl.DataSource = dt;
              }
           
        }

        DataTable worknoBadclassDetail(string workno)
        {
            DataTable dt = null;
            string sql = "  select  org_id 组织,workno 工单号,materialcode 编码,serialNo 序列号,defect 不良现象,setno 不良位号,TQTY 不良数量,materialName 描述,checkDate 时间  from   ";
            sql += "  (   select 'SHL' org_id,a.workno ,a.illMcode materialcode,a.serialNo ,c.defectname defect,a.illPosition setno,1 TQTY,illMdescript materialName,illDate checkDate  from OEM_Repair a left join DefectCode c on a.illReason =c.defectcode   where c.deptid = 'SMT' and  a.type='炉后和AOI检查'  and  illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00'  and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59' ";
            sql += "   union all   select  org_id,workno,materialcode,serialNo,defect,setno,TQTY,materialName,checkDate from SMT_WorknoRepair  where  station in ('AOI检测','炉后QC检测') and checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'    ";
            sql += "     ) i  "+workno+ "  ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable modelBadclassDetail(string model)
        {
            DataTable dt = null;
            string sql = "  select  org_id 组织,workno 工单号,materialcode 编码,serialNo 序列号,defect 不良现象,setno 不良位号,TQTY 不良数量,materialName 描述,checkDate 时间  from   ";
            sql += "  (   select 'SHL' org_id,a.workno ,a.illMcode materialcode,a.serialNo ,c.defectname defect,a.illPosition setno,1 TQTY,illMdescript materialName,illDate checkDate  from OEM_Repair a left join DefectCode c on a.illReason =c.defectcode   where  c.deptid = 'SMT' and  a.type='炉后和AOI检查'  and  illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00'  and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59' ";
            sql += "   union all   select  org_id,workno,materialcode,serialNo,defect,setno,TQTY,materialName,checkDate from SMT_WorknoRepair  where  station in ('AOI检测','炉后QC检测') and checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'    ";
            sql += "     ) i  where materialcode = '"+ model+ "'  ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable customerBadclassDetail(string customer)
        {
            DataTable dt = null;
            string sql = "  select  i.org_id 组织,i.workno 工单号,materialcode 编码,o.customer 客户,o.laName 线别,serialNo 序列号,defect 不良现象,setno 不良位号,TQTY 不良数量,materialName 描述,checkDate 时间  from    ";
            sql += "  (  select 'SHL' org_id,a.workno ,a.illMcode materialcode,a.serialNo ,c.defectname defect,a.illPosition setno,1 TQTY,illMdescript materialName,illDate checkDate  from OEM_Repair a left join DefectCode c on a.illReason =c.defectcode   where c.deptid = 'SMT' and  a.type='炉后和AOI检查'  and  illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00'  and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  ";
            sql += "  union all   select  org_id,workno,materialcode,serialNo,defect,setno,TQTY,materialName,checkDate from SMT_WorknoRepair  where  station in ('AOI检测','炉后QC检测') and checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'  ";
            sql += "   ) i  left join   (  select org_id,workno,customer,l.laName from  OEM_Prod e left join latype l on e.latype = l.la     ";
            sql += "   )  o on i.org_id = o.org_id and i.workno = o.workno where o.customer in("+customer+ ") ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }

        DataTable latypeBadclassDetail(string latype)
        {
            DataTable dt = null;
            string sql = "  select  i.org_id 组织,i.workno 工单号,materialcode 编码,o.laName 线别,serialNo 序列号,defect 不良现象,setno 不良位号,TQTY 不良数量,materialName 描述,checkDate 时间  from  ";
            sql += "   (  select 'SHL' org_id,a.workno ,a.illMcode materialcode,a.serialNo ,c.defectname defect,a.illPosition setno,1 TQTY,illMdescript materialName,illDate checkDate  from OEM_Repair a left join DefectCode c on a.illReason =c.defectcode   where c.deptid = 'SMT' and  a.type='炉后和AOI检查'  and  illDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00'  and illDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'     ";
            sql += "   union all   select  org_id,workno,materialcode,serialNo,defect,setno,TQTY,materialName,checkDate from SMT_WorknoRepair  where  station in ('AOI检测','炉后QC检测') and checkDate >='" + txtstartdate.DateTime.ToString("yyyy-MM-dd") + " 00:00:00' and checkDate <='" + txtenddate.DateTime.ToString("yyyy-MM-dd") + " 23:59:59'      ";
            sql += "  ) i  inner  join   (   select org_id,workno,l.laName from    (select org_id, workno,latype from   ( select  org_id,workno,latype from OEM_Prod   union all  select  org_id,workno,latype from SMT_Prod   ";
            sql += "    ) w  where w.latype <>'' and w.latype is not null  group by org_id,workno,latype ) s left join latype l on s.latype = l.la  where l.laName in("+ latype+ ") ";
            sql += "   ) o on i.org_id = o.org_id and i.workno = o.workno  ";
            dt = DbAccess.SelectBySql(sql).Tables[0];
            return dt;
        }
        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            string name = "SMT不良现象清单";
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

        private void sBtnExportBadClass_Click(object sender, EventArgs e)
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

    }
}