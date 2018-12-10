namespace DX_QMS
{
    partial class QAPackCheck
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.databind = new DevExpress.XtraGrid.GridControl();
            this.gridView = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.gridColumn1 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn2 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.gridColumn3 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.btnsearch = new DevExpress.XtraEditors.SimpleButton();
            this.lblinfo = new DevExpress.XtraEditors.LabelControl();
            this.txtworkno = new DevExpress.XtraEditors.TextEdit();
            this.txtboxsn = new DevExpress.XtraEditors.TextEdit();
            this.txtcolorsn = new DevExpress.XtraEditors.TextEdit();
            this.txtsn = new DevExpress.XtraEditors.TextEdit();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.lblsn = new DevExpress.XtraLayout.LayoutControlItem();
            this.lblcolorsn = new DevExpress.XtraLayout.LayoutControlItem();
            this.lblboxsn = new DevExpress.XtraLayout.LayoutControlItem();
            this.lblworkno = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.databind)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtworkno.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtboxsn.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtcolorsn.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtsn.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblsn)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblcolorsn)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblboxsn)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblworkno)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.layoutControl1.Controls.Add(this.databind);
            this.layoutControl1.Controls.Add(this.btnsearch);
            this.layoutControl1.Controls.Add(this.lblinfo);
            this.layoutControl1.Controls.Add(this.txtworkno);
            this.layoutControl1.Controls.Add(this.txtboxsn);
            this.layoutControl1.Controls.Add(this.txtcolorsn);
            this.layoutControl1.Controls.Add(this.txtsn);
            this.layoutControl1.Location = new System.Drawing.Point(60, 31);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.layoutControlGroup1;
            this.layoutControl1.Size = new System.Drawing.Size(647, 459);
            this.layoutControl1.TabIndex = 17;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // databind
            // 
            this.databind.Location = new System.Drawing.Point(12, 108);
            this.databind.MainView = this.gridView;
            this.databind.Name = "databind";
            this.databind.Size = new System.Drawing.Size(623, 339);
            this.databind.TabIndex = 19;
            this.databind.UseEmbeddedNavigator = true;
            this.databind.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView});
            // 
            // gridView
            // 
            this.gridView.Appearance.HeaderPanel.Options.UseTextOptions = true;
            this.gridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridView.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.gridColumn1,
            this.gridColumn2,
            this.gridColumn3});
            this.gridView.GridControl = this.databind;
            this.gridView.Name = "gridView";
            this.gridView.OptionsView.ShowGroupPanel = false;
            // 
            // gridColumn1
            // 
            this.gridColumn1.AppearanceCell.Options.UseTextOptions = true;
            this.gridColumn1.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridColumn1.Caption = "机身号";
            this.gridColumn1.FieldName = "SN";
            this.gridColumn1.Name = "gridColumn1";
            this.gridColumn1.Visible = true;
            this.gridColumn1.VisibleIndex = 0;
            // 
            // gridColumn2
            // 
            this.gridColumn2.AppearanceCell.Options.UseTextOptions = true;
            this.gridColumn2.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridColumn2.Caption = "彩盒号";
            this.gridColumn2.FieldName = "CorSN";
            this.gridColumn2.Name = "gridColumn2";
            this.gridColumn2.Visible = true;
            this.gridColumn2.VisibleIndex = 1;
            // 
            // gridColumn3
            // 
            this.gridColumn3.AppearanceCell.Options.UseTextOptions = true;
            this.gridColumn3.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridColumn3.Caption = "外箱号";
            this.gridColumn3.FieldName = "LotNo";
            this.gridColumn3.Name = "gridColumn3";
            this.gridColumn3.Visible = true;
            this.gridColumn3.VisibleIndex = 2;
            // 
            // btnsearch
            // 
            this.btnsearch.Location = new System.Drawing.Point(280, 78);
            this.btnsearch.MaximumSize = new System.Drawing.Size(80, 35);
            this.btnsearch.Name = "btnsearch";
            this.btnsearch.Size = new System.Drawing.Size(80, 26);
            this.btnsearch.StyleController = this.layoutControl1;
            this.btnsearch.TabIndex = 9;
            this.btnsearch.Text = "查询";
            this.btnsearch.Click += new System.EventHandler(this.btnsearch_Click);
            // 
            // lblinfo
            // 
            this.lblinfo.Location = new System.Drawing.Point(12, 60);
            this.lblinfo.Name = "lblinfo";
            this.lblinfo.Size = new System.Drawing.Size(623, 14);
            this.lblinfo.StyleController = this.layoutControl1;
            this.lblinfo.TabIndex = 8;
            // 
            // txtworkno
            // 
            this.txtworkno.Location = new System.Drawing.Point(365, 36);
            this.txtworkno.Name = "txtworkno";
            this.txtworkno.Size = new System.Drawing.Size(270, 20);
            this.txtworkno.StyleController = this.layoutControl1;
            this.txtworkno.TabIndex = 7;
            // 
            // txtboxsn
            // 
            this.txtboxsn.Location = new System.Drawing.Point(51, 36);
            this.txtboxsn.Name = "txtboxsn";
            this.txtboxsn.Size = new System.Drawing.Size(271, 20);
            this.txtboxsn.StyleController = this.layoutControl1;
            this.txtboxsn.TabIndex = 6;
            this.txtboxsn.KeyUp += new System.Windows.Forms.KeyEventHandler(this.txtboxsn_KeyUp);
            // 
            // txtcolorsn
            // 
            this.txtcolorsn.Location = new System.Drawing.Point(365, 12);
            this.txtcolorsn.Name = "txtcolorsn";
            this.txtcolorsn.Size = new System.Drawing.Size(270, 20);
            this.txtcolorsn.StyleController = this.layoutControl1;
            this.txtcolorsn.TabIndex = 5;
            this.txtcolorsn.KeyUp += new System.Windows.Forms.KeyEventHandler(this.txtcolorsn_KeyUp);
            // 
            // txtsn
            // 
            this.txtsn.Location = new System.Drawing.Point(51, 12);
            this.txtsn.Name = "txtsn";
            this.txtsn.Size = new System.Drawing.Size(271, 20);
            this.txtsn.StyleController = this.layoutControl1;
            this.txtsn.TabIndex = 4;
            this.txtsn.KeyUp += new System.Windows.Forms.KeyEventHandler(this.txtsn_KeyUp);
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.layoutControlGroup1.GroupBordersVisible = false;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.lblsn,
            this.lblcolorsn,
            this.lblboxsn,
            this.lblworkno,
            this.layoutControlItem5,
            this.layoutControlItem6,
            this.emptySpaceItem2,
            this.emptySpaceItem1,
            this.layoutControlItem2});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(647, 459);
            this.layoutControlGroup1.TextVisible = false;
            // 
            // lblsn
            // 
            this.lblsn.Control = this.txtsn;
            this.lblsn.Location = new System.Drawing.Point(0, 0);
            this.lblsn.Name = "lblsn";
            this.lblsn.Size = new System.Drawing.Size(314, 24);
            this.lblsn.Text = "机身号";
            this.lblsn.TextSize = new System.Drawing.Size(36, 14);
            // 
            // lblcolorsn
            // 
            this.lblcolorsn.Control = this.txtcolorsn;
            this.lblcolorsn.Location = new System.Drawing.Point(314, 0);
            this.lblcolorsn.Name = "lblcolorsn";
            this.lblcolorsn.Size = new System.Drawing.Size(313, 24);
            this.lblcolorsn.Text = "彩盒号";
            this.lblcolorsn.TextSize = new System.Drawing.Size(36, 14);
            // 
            // lblboxsn
            // 
            this.lblboxsn.Control = this.txtboxsn;
            this.lblboxsn.Location = new System.Drawing.Point(0, 24);
            this.lblboxsn.Name = "lblboxsn";
            this.lblboxsn.Size = new System.Drawing.Size(314, 24);
            this.lblboxsn.Text = "外箱号";
            this.lblboxsn.TextSize = new System.Drawing.Size(36, 14);
            // 
            // lblworkno
            // 
            this.lblworkno.Control = this.txtworkno;
            this.lblworkno.Location = new System.Drawing.Point(314, 24);
            this.lblworkno.Name = "lblworkno";
            this.lblworkno.Size = new System.Drawing.Size(313, 24);
            this.lblworkno.Text = "工单号";
            this.lblworkno.TextSize = new System.Drawing.Size(36, 14);
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.lblinfo;
            this.layoutControlItem5.Location = new System.Drawing.Point(0, 48);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(627, 18);
            this.layoutControlItem5.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem5.TextVisible = false;
            // 
            // layoutControlItem6
            // 
            this.layoutControlItem6.Control = this.btnsearch;
            this.layoutControlItem6.Location = new System.Drawing.Point(268, 66);
            this.layoutControlItem6.Name = "layoutControlItem6";
            this.layoutControlItem6.Size = new System.Drawing.Size(84, 30);
            this.layoutControlItem6.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem6.TextVisible = false;
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(0, 66);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(268, 30);
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(352, 66);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(275, 30);
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.databind;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 96);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(627, 343);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // QAPackCheck
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(772, 498);
            this.Controls.Add(this.layoutControl1);
            this.Name = "QAPackCheck";
            this.Text = "QA核对";
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.databind)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtworkno.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtboxsn.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtcolorsn.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtsn.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblsn)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblcolorsn)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblboxsn)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblworkno)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraEditors.SimpleButton btnsearch;
        private DevExpress.XtraEditors.LabelControl lblinfo;
        private DevExpress.XtraEditors.TextEdit txtworkno;
        private DevExpress.XtraEditors.TextEdit txtboxsn;
        private DevExpress.XtraEditors.TextEdit txtcolorsn;
        private DevExpress.XtraEditors.TextEdit txtsn;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem lblsn;
        private DevExpress.XtraLayout.LayoutControlItem lblcolorsn;
        private DevExpress.XtraLayout.LayoutControlItem lblboxsn;
        private DevExpress.XtraLayout.LayoutControlItem lblworkno;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraGrid.GridControl databind;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn1;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn2;
        private DevExpress.XtraGrid.Columns.GridColumn gridColumn3;
    }
}