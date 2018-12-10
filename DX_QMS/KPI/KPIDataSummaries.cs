using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.Utils;
using DevExpress.Data;
using DevExpress.XtraGrid;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Menu;
using DevExpress.Utils.Menu;
using System.Data.SqlClient;
using DX_QMS.Common;

namespace DevExpress.XtraGrid.Demos {
    /// <summary>
    /// Summary description for DataSummaries.
    /// </summary>
    public partial class DataSummaries : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        DXPopupMenu formatRulesMenu = new DXPopupMenu();
        DataSet ds = new DataSet();
        public DataSummaries()
        {
            //
            // Required for Windows Form Designer support
            //
           
            InitializeComponent();

            InitMDBData();
           //////////// InitSummaries();
            InitEditing();
        }

        private GridGroupSummaryItemCollection gsiSummary, gsiMultiSummary, gsiMultiSummaryDetail, gsiDisplaySummary, gsiDisplaySummaryDetails;


        private bool displaySummary = false;
        #region Init
        public DevExpress.XtraGrid.Views.Base.BaseView ExportView { get { return advBandedGridView; } }


        private GridControl CurrentGrid { get { return gridControl; } }
        private GridView CurrentView { get { return advBandedGridView; } }
        private GridView CurrentDetailView { get { return gridView; } }


        private void sBtnselect_Click(object sender, EventArgs e)
        {




        }




        protected void InitMDBData()
        {
            SqlConnection conText = new SqlConnection(DbAccess.connSql);

            try
            {

                //     conText.Open();

                //     string adsql = @"  select KpiYeartime+'Äê' KpiYeartime,businessType ,indicatorsName,checkDepartment ,undertakeLevel ,qualityMan ,dataSource from QMS_KPIDataSource 		   
                //group by KpiYeartime, businessType ,indicatorsName,checkDepartment ,undertakeLevel,qualityMan ,dataSource ";
                //     SqlDataAdapter adda = new SqlDataAdapter(adsql, conText);
                //     adda.Fill(ds, "adtable");

                //     string desql = @"  select kpiMonthtime,indicatorsName ,lastYearkpiData ,lastYearMonthkpiData ,thisYearkpiData ,kpiData ,ifComplete ,mathFormula ,Remark 
                //                  from QMS_KPIDataSource  ";
                //     SqlDataAdapter deda = new SqlDataAdapter(desql, conText);
                //     deda.Fill(ds, "detable");

                conText.Open();

                string adsql = @"  select KpiYeartime+'Äê' KpiYeartime,businessType ,indicatorsName,checkDepartment ,undertakeLevel ,qualityMan ,dataSource from QMS_KPIDataSource 		   
		         group by KpiYeartime, businessType ,indicatorsName,checkDepartment ,undertakeLevel,qualityMan ,dataSource ";
                SqlDataAdapter adda = new SqlDataAdapter(adsql, conText);
                adda.Fill(ds, "adtable");

                string desql = @"  select kpiMonthtime,indicatorsName ,lastYearkpiData ,lastYearMonthkpiData ,thisYearkpiData ,kpiData ,ifComplete ,mathFormula ,Remark 
		                           from QMS_KPIDataSource  ";
                adda = new SqlDataAdapter(desql, conText);
                adda.Fill(ds, "detable");
            }
            catch
            {
                conText.Close();
            }

            ds.Relations.Add("adtable",
                ds.Tables["adtable"].Columns["indicatorsName"],
                ds.Tables["detable"].Columns["indicatorsName"], false);

          
            var node = new GridLevelNode();
            node.LevelTemplate = gridView;
            node.RelationName = "adtable";
            gridControl.LevelTree.Nodes.AddRange(new GridLevelNode[]
            {
                node
            });

            gridControl.DataSource = ds.Tables["adtable"];


        }
        private void InitEditing()
        {
           
            gridControl.ForceInitialize();
            /*
            colSupplierID.GroupIndex = 0;
            colPrice.GroupIndex = 0;
            SetShowFooter(chShowFooter.Checked);
            chAlignSummary.Checked = true;
            OnSummaryChecked(chDisplaySummary);
            */
        }
        private void UpdateMasterDetailSettings()
        {
            CurrentView.BeginUpdate();
            CurrentView.ExpandAllGroups();
            CurrentView.FocusedRowHandle = 0;
            CurrentView.TopRowIndex = 0;
            CurrentView.SetMasterRowExpanded(CurrentView.FocusedRowHandle, true);
            if (ceMasterDetail.Checked)
            {
                if (FirstDetailView != null) FirstDetailView.ExpandGroupRow(-1);
            }
            else
                CurrentDetailView.ExpandAllGroups();
            CurrentView.EndUpdate();
        }
        GridView FirstDetailView
        { 
            get {
                return CurrentView.GetDetailView(0, 0) as GridView;
                }
        }
        //<gridControl1>
        private void InitSummaries()
        {
            /*
            //~row summary 
            gsiSummary = new GridGroupSummaryItemCollection(CurrentView);
            gsiSummary.Add(SummaryItemType.Count, "ProductID");

            // ~multi row summary 
            gsiMultiSummary = new GridGroupSummaryItemCollection(CurrentView);
            gsiMultiSummary.Add(SummaryItemType.Count, "ProductID");
            gsiMultiSummary.Add(SummaryItemType.Average, "UnitPrice", null);

           //  ~multi row summary for detail
            gsiMultiSummaryDetail = new GridGroupSummaryItemCollection(CurrentDetailView);
            gsiMultiSummaryDetail.Add(SummaryItemType.Count, "OrderID");
            gsiMultiSummaryDetail.Add(SummaryItemType.Sum, "SubTotal", null);

            // ~row footer summary 
            gsiDisplaySummary = new GridGroupSummaryItemCollection(CurrentView);
            gsiDisplaySummary.Add(SummaryItemType.Max, "UnitsOnOrder", colUnitsOnOrder);
            gsiDisplaySummary.Add(SummaryItemType.Sum, "UnitsInStock", colUnitsInStock);
            gsiDisplaySummary.Add(SummaryItemType.Average, "UnitPrice", colUnitPrice);
            gsiDisplaySummary.Add(SummaryItemType.Count, "ProductName", colProductName);

            // ~row footer summary for details 
            gsiDisplaySummaryDetails = new GridGroupSummaryItemCollection(CurrentDetailView);
            gsiDisplaySummaryDetails.Add(SummaryItemType.Sum, "SubTotal", colSubTotal);
            gsiDisplaySummaryDetails.Add(SummaryItemType.Min, "Quantity", colQuantity);

            */

        }
        //</gridControl1>

        private void DataSummaries_Load(object sender, System.EventArgs e) {
            UpdateMasterDetailSettings();
        }
        #endregion
        #region Editing
        //<chShowFooter>
        private void SetShowFooter(bool show) {
            CurrentView.OptionsView.ShowFooter = show;
            CurrentDetailView.OptionsView.ShowFooter = show;
        }
        private void chShowFooter_CheckedChanged(object sender, System.EventArgs e)
        {
            CheckEdit chb = sender as CheckEdit;
            SetShowFooter(chb.Checked);
            
        }
       

        private void SaveDisplaySummary() {
            if (displaySummary) {
                gsiDisplaySummary.Assign(CurrentView.GroupSummary);
                gsiDisplaySummaryDetails.Assign(CurrentDetailView.GroupSummary);
                CurrentView.OptionsView.GroupFooterShowMode = GroupFooterShowMode.VisibleIfExpanded;
                CurrentDetailView.OptionsView.GroupFooterShowMode = GroupFooterShowMode.VisibleIfExpanded;
            }
            displaySummary = false;
        }
        #endregion
        decimal GetSubTotalFromDataRow(DataRow row)
        {
            decimal q = Convert.ToDecimal(row["Quantity"]);
            decimal p = Convert.ToDecimal(row["UnitPrice"]);
            decimal d = Convert.ToDecimal(row["Discount"]);
            return q * p * (1 - d);
        }

        private void gridView1_CustomUnboundColumnData(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs e)
        {
            if (e.IsSetData || e.Column.FieldName != "SubTotal") return;
            GridView view = sender as GridView;
            e.Value = GetSubTotalFromDataRow(((DataRowView)e.Row).Row);
        }

        bool updateInfo = false;
        private void chSummary_CheckedChanged(object sender, EventArgs e)
        {
            if (updateInfo) return;
           ///////// OnSummaryChecked(sender as CheckEdit);
        }

    
        private void OnSummaryChecked(CheckEdit edit)
        {
            if(edit.Properties.Tag == null) return;
            updateInfo = true;
            string caption = edit.Properties.Tag.ToString();
            switch (caption) {
                case "Summary":
                    chSummary.Checked = true;
                    chAlignSummary.Enabled = false;
                    SaveDisplaySummary();
                    CurrentView.GroupSummary.Assign(gsiSummary);
                    CurrentDetailView.GroupSummary.Assign(gsiSummary);
                    break;
                case "Multi Summary":
                    chMultiSummary.Checked = true;
                    chAlignSummary.Enabled = false;
                    SaveDisplaySummary();
                    CurrentView.GroupSummary.Assign(gsiMultiSummary);
                    CurrentDetailView.GroupSummary.Assign(gsiMultiSummaryDetail);
                    break;
                case "Display Summary":
                    chDisplaySummary.Checked = true;
                    chAlignSummary.Enabled = true;
                    displaySummary = true;
                    CurrentView.OptionsView.GroupFooterShowMode = GroupFooterShowMode.VisibleAlways;
                    CurrentDetailView.OptionsView.GroupFooterShowMode = GroupFooterShowMode.VisibleAlways;
                    CurrentView.GroupSummary.Assign(gsiDisplaySummary);
                    CurrentDetailView.GroupSummary.Assign(gsiDisplaySummaryDetails);
                    break;
            }
            UpdateAlignSummary();
            UpdateMasterDetailSettings();
            updateInfo = false;
        }
        void UpdateAlignSummary() {
            CurrentDetailView.OptionsBehavior.AlignGroupSummaryInGroupRow = 
                chAlignSummary.Enabled && chAlignSummary.Checked ? DefaultBoolean.True : DefaultBoolean.False;
        }
     
        void chAlignSummary_CheckedChanged(object sender, EventArgs e)
        {
            UpdateAlignSummary();
        }

        private void ceMasterDetail_CheckedChanged(object sender, EventArgs e)
        {

            /*
            if(ceMasterDetail.Checked) {
                colOrderID.GroupIndex = -1;
                colPrice.GroupIndex = 0;
                gridControl1.MainView = CurrentView;
                ////gridControl1.DataSource = dsNWindProducts1.Products;

                colProduct.Visible = false;
            } else {
                colPrice.GroupIndex = -1;
                colOrderID.GroupIndex = 0;
                gridControl1.MainView = CurrentDetailView;
                ////gridControl1.DataSource = dsNWindProducts1.Order_Details;


                colProduct.VisibleIndex = 0;
            }
            UpdateMasterDetailSettings();
            */
             

        }

        private void gridView_RowCellClick(object sender, RowCellClickEventArgs e)
        {

            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                GridFormatRuleMenuItems items = new GridFormatRuleMenuItems(gridView, e.Column, formatRulesMenu.Items);
                if (items.Count > 0)
                    MenuManagerHelper.ShowMenu(formatRulesMenu, gridControl.LookAndFeel, gridControl.MenuManager, gridControl, new Point(e.X, e.Y));
            }

        }


        public bool AllowGenerateReport { get { return false; } }
       
    }
}
