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
using DevExpress.XtraEditors;
using DevExpress.XtraTreeList.Nodes;
using System.Data;
using System.Data.SqlClient;
using DX_QMS.Common;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Base;

namespace DX_QMS
{
    public partial class ExceptionDetails : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public ExceptionDetails()
        {
            InitializeComponent();
            gridView.PrintExportProgress += new ProgressChangedEventHandler(gridView_PrintExportProgress);
        }

        public delegate bool MethodCaller();

        DataTable BadSituation()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("不良类别", System.Type.GetType("System.String"));
            dt.Columns.Add("不良现象", System.Type.GetType("System.String"));
            string[] bad = { "包装", "标签", "数量", "错料", "外观", "结构尺寸", "试装不良", "功能", "环保", "其它" };
            string[] situation = 
                {
                "散料、漏气、不符合ESD、MSL包装要求",
                 "五要素缺失、超期、批次、版本、型号",
                 "叉板超标、多装、短装、打叉板未减数、少配件等不良情况",
                 "与BOM描述不一致、混料",
                 "划痕、丝印、色差、引脚、打点、漏工序等",
                 "长宽厚、平面度、翘曲",
                 "螺孔打不到底，造成滑牙，或者装配间隙大，无法匹配，造成卡塞或者装不进去",
                 "可靠性测试（丝印粘着力测试、低温跌落实验，48H盐雾等）、阻容感值、发光测试",
                 "RoHS",
                 "材质问题或其他出现问题"
                 };
            for (int i = 0; i < 10; i++)
            {
                dt.Rows.Add(bad[i],situation[i]);
            }

            return dt;
        }
        void brand(string materialcategory)
        {
            TreeNode tmp = new TreeNode();
            tmp.Text = materialcategory;
            tmp.Name = materialcategory;
            /*treeView1.Nodes.Add(tmp);*/
            string[] MaterialCategory = { "主营","EMS"};
            string[] Material_zhu = { "配件原材料","电子料","结构件","包材","辅料","外购成品", "其他" };
            string[] Matetial_EMS = { "电子料","结构件","包材","辅料","其它"};
            string[] Category_zhu_jie = { "开关","电子结构","连接器","线材","特殊结构件","螺钉","螺母","小五金","橡硅胶","大五金","铝壳"};
            string[] Category_zhu_dian = { "PCB/FCB","电容","电阻","电感","二极管","三极管","场效应管","集成电路","晶振(石英)","滤波器","鉴频器/解调器","保险类"};
            string[] Category_EMS_jie = { "PCB/FCB","电容","电阻","电感","二极管","三极管","场效应管","IC","晶振(石英)" };
            string[] Category_EMS_dian ={ "开关类","连接器","小五金类","线材"};
            if (tmp.Name == "主营")
                {
                    for (int n =0; n < Material_zhu.Length;n++ )
                    {
                        TreeNode node_zhu = new TreeNode();
                        node_zhu.Text = Material_zhu[n];
                        node_zhu.Name = Material_zhu[n];
                        tmp.Nodes.Add(node_zhu);
                        if (node_zhu.Name == "结构件")
                        {
                            for (int m = 0; m < Category_zhu_jie.Length; m++)
                            {
                                TreeNode jiegouti = new TreeNode();
                                jiegouti.Text = Category_zhu_jie[m];
                                jiegouti.Name = Category_zhu_jie[m];
                                node_zhu.Nodes.Add(jiegouti);
                            }                          
                        }
                        else if (node_zhu.Name =="电子料")
                        {
                            for (int m = 0; m < Category_zhu_dian.Length; m++)
                            {
                                TreeNode dianzilaio = new TreeNode();
                                dianzilaio.Text = Category_zhu_dian[m];
                                dianzilaio.Name = Category_zhu_dian[m];
                                node_zhu.Nodes.Add(dianzilaio);

                            }
                        }
                    }
                }
            if (tmp.Name == "EMS")
                 {
                    for (int n = 0; n < Matetial_EMS.Length; n++)
                    {
                        TreeNode node_EMS = new TreeNode();
                        node_EMS.Text = Matetial_EMS[n];
                        node_EMS.Name = Matetial_EMS[n];
                        tmp.Nodes.Add(node_EMS);
                        if (node_EMS.Name == "结构件")
                        {
                            for (int m = 0; m < Category_EMS_jie.Length; m++)
                            {
                                TreeNode jiegoujian = new TreeNode();
                                jiegoujian.Text = Category_EMS_jie[m];
                                jiegoujian.Name = Category_EMS_jie[m];
                                node_EMS.Nodes.Add(jiegoujian);
                            }

                        }
                        else if (node_EMS.Name == "电子料")
                        {
                            for (int m = 0; m < Category_EMS_dian.Length; m++)
                            {
                                TreeNode dianziliao = new TreeNode();
                                dianziliao.Text = Category_EMS_dian[m];
                                dianziliao.Name = Category_EMS_dian[m];
                                node_EMS.Nodes.Add(dianziliao);
                            }
                        }
                    }
                 }
        }


        private void ExceptionDetails_Load(object sender, EventArgs e)
        {
            //brand("主营");
            //brand("EMS");
            Exception_category.Properties.DataSource = BadSituation();
            Exception_category.Properties.DisplayMember = "不良类别";
            Exception_category.Properties.ValueMember = "不良类别";
            Btn_reset_Click(sender, e);
            InitExportData();
            date_start.Text = DateTime.Now.ToString("yyyy-MM-dd");
            date_end.Text = DateTime.Now.ToString("yyyy-MM-dd");
        }
        string[,] exportData = new string[,] {
            {"Microsoft Excel 2007 Document", "Microsoft Excel|*.xlsx", "xlsx"},
            {"Microsoft Excel Document", "Microsoft Excel|*.xls", "xls"},
            {"PDF Document", "PDF Files|*.pdf", "pdf"},
            {"Text Document", "Text Files|*.txt", "txt"}};
        void InitExportData()
        {
            for (int i = 0; i < exportData.GetLength(0); i++)
                cbExport.Properties.Items.Add(exportData.GetValue(i, 0));
            cbExport.SelectedIndex = 0;
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


        /*
        private void AsyncShowDetail(IAsyncResult result)
        {
            MethodCaller aysnDelegate = result.AsyncState as MethodCaller;
            if (aysnDelegate != null)
            {
                bool success = aysnDelegate.EndInvoke(result);
                if (success)
                {
                    ConnDB conn = new ConnDB();
                    string strsql;
                    strsql = "";
                    DataSet ds = conn.ReturnDataSet(strsql);

                    Action<DataSet> action = (data) =>
                    {
                        gridControl1.DataSource = data.Tables[0].DefaultView;
                        gridView1.Columns[1].OptionsColumn.ReadOnly = true;
                        simpleButtonImport.Enabled = true;
                        simpleButtonClear.Enabled = true;
                    };
                    Invoke(action, ds);
                    conn.Close();
                }
            }
        }

        private bool Import()
        {
           
       }

      */

        private void Btn_search_Click(object sender, EventArgs e)
        {
            string where = " where t.TestFinalResult = '拒收' and t.TestResult = 'NG' and t.TestItem not like '_值测试' ";
            string bdate = date_start.DateTime.ToString("yyyy-MM-dd") ;
            string edate = date_end.DateTime.ToString("yyyy-MM-dd");
            string exceptioncategory = Exception_category.Text;
            string materialcatagory = Material_category.Text;
            string productcode = txtproductcode.Text.Trim();
            string supplier = txtsupplier.Text.Trim();
            //DataSet ds = null;
           // ds = SelectTestListRecord("异常明细表",exceptioncategory, materialcatagory, "", "", "", "", txtproductcode.Text.Trim(), "", "", txtsupplier.Text.Trim(), bdate, edate);

            //ds = SelectTestListRecord("异常明细表","", "主营：配件原材料", "", "", "", "", "", "", "","","", "");

            if (!string.IsNullOrEmpty(supplier))
            {
                where += " and vendorname like '%" + supplier + "%' ";
            }
            if (!string.IsNullOrEmpty(productcode))
            {
                where += " and d.materialcode like '%" + productcode + "%' ";
            }
            if (!string.IsNullOrEmpty(bdate))
            {
                where += " and t.testtime>= '" + bdate+ " 00:00:00" + "' ";
            }
            if (!string.IsNullOrEmpty(edate))
            {
                where += " and t.testtime<='" + edate + " 23:59:59" + "' ";
            }
            if (!string.IsNullOrEmpty(materialcatagory))
            {
                where += " and  t.receptid  = '"+materialcatagory+ "' ";
            }
            if (!string.IsNullOrEmpty(exceptioncategory))
            {
                where += " and NGtype ='" + exceptioncategory + "' ";
            }


            string sql = @" select d.deliveryid 接收单号,t.Productcode 物料编码,t.LotNo 批次号,s.supersort+':' +s.subsort  物料类别,m.materialname 物料描述,
	                 d.vendorname 供应商或客户,d.vendorcode 供应商或客户代码,t.TestItem 测试项目,t.TestSubItem 测试子项目,t.Qty 总数,t.Testqty 检验数量,t.Remarks 备注信息,t.DCode DC,t.Mdate 生产日期,t.ExpiryDate 失效日期,
		             t.NGtype 不良类别,t.TestUser 检验员,t.testtime 检验时间
	                 from IQC_TestList t left join delivery  d  on t.Productcode = d.materialcode and t.receptid = d.deliveryid left join MaterialSpec m on t.Productcode = m.materialcode 
	                 left join materialsort s on t.Productcode = s.sortcode  ";
            sql += where + " order by t.TestTime desc ";

            DataTable dt = DbAccess.SelectBySql(sql).Tables[0];

            if (dt !=null  && dt.Rows.Count > 0)
            {
                gridList.DataSource = dt;
            }
            else
            {
                MessageBox.Show("没有符合条件的记录");
                gridList.DataSource = null;
            }

            //if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            //{
            //    gridList.DataSource = ds.Tables[0];
            //}
            //else
            //{
            //    MessageBox.Show("没有符合条件的记录");
            //    gridList.DataSource = null;
            //}
        }
        private void Material_category_QueryPopUp(object sender, CancelEventArgs e)
        {
            //PopupContainerEdit popupedit = (PopupContainerEdit)sender;
            //popupContainerControl1.Width = popupedit.Width;
        }
        string findparent(TreeNode node)
        {
            if ( (node.Parent.Name == "主营") || (node.Parent.Name == "EMS" ))
            {
                return node.Parent.Name;
            }
            else
            {
               return findparent(node.Parent);
             }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {           
            //if(findparent(treeView1.SelectedNode) == "主营")
            //  Material_category.EditValue = findparent(treeView1.SelectedNode) + "：" + treeView1.SelectedNode.Name;
            //if (findparent(treeView1.SelectedNode) == "EMS")
            //  Material_category.EditValue = findparent(treeView1.SelectedNode) + "：" + treeView1.SelectedNode.Name;
            //  Material_category.ClosePopup();
        }
        private void Btn_reset_Click(object sender, EventArgs e)
        {
            txtsupplier.Text = "";
            txtsupplier.EditValue = null;
            txtproductcode.Text = "";
            txtproductcode.EditValue = null;
            date_start.Text = "";
            date_start.EditValue = null;
            date_end.Text = "";
            date_end.EditValue = null;
            Material_category.Text = "";
            Material_category.EditValue = null;
            Exception_category.Text = "";
            Exception_category.EditValue = null;
        }

        private void gridView_CustomSummaryCalculate(object sender, DevExpress.Data.CustomSummaryEventArgs e)
        {
            int customSum = 0;
            int summaryId = Convert.ToInt32((e.Item as DevExpress.XtraGrid.GridSummaryItem).Tag);
            if (e.SummaryProcess == DevExpress.Data.CustomSummaryProcess.Start)
            {
                customSum = 1;
            }
            if (((DevExpress.XtraGrid.GridSummaryItem)e.Item).FieldName == "批次号")
            {
                //for (int i = 0; i < gridView.RowCount; i++)
                //{
                //    string temp = gridView.GetDataRow(i)["批次号"].ToString();
                //    for (int j = i + 1; j < gridView.RowCount; j++)
                //    {
                //        if (temp == gridView.GetDataRow(j)["批次号"].ToString())
                //            break;
                //        customSum++;
                //    }
                //}
               // e.TotalValue = customSum;
                

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
        private string ShowSaveFileDialog(string title, string filter)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            /////// string name = Application.ProductName;
            string name = "来料异常明细";
            int n = name.LastIndexOf(".") + 1;
            if (n > 0) name = name.Substring(n, name.Length - n);
            dlg.Title = "导出...... " + title;
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
            if (XtraMessageBox.Show("你要打开此文件?", "导出...", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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

        private void sbExport_Click(object sender, EventArgs e)
        {
            int index = cbExport.SelectedIndex;
            if (index < 0) return;
            string fileName = ShowSaveFileDialog(exportData.GetValue(index, 0).ToString(), exportData.GetValue(index, 1).ToString());
            if (fileName == string.Empty) return;
            ExportToEx(fileName, exportData.GetValue(index, 2).ToString(), gridView);
            OpenFile(fileName);
        }
    }
}