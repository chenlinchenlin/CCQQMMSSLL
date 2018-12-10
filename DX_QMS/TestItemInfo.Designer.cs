namespace DX_QMS
{
    partial class TestItemInfo
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
            this.txtid = new System.Windows.Forms.TextBox();
            this.cbTestType = new System.Windows.Forms.ComboBox();
            this.txtTestItem = new System.Windows.Forms.TextBox();
            this.btnReset = new DevExpress.XtraEditors.SimpleButton();
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.dgvDefect = new DevExpress.XtraGrid.GridControl();
            this.gridView = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.TestType = new DevExpress.XtraGrid.Columns.GridColumn();
            this.TestItem = new DevExpress.XtraGrid.Columns.GridColumn();
            this.Item = new DevExpress.XtraGrid.Columns.GridColumn();
            this.btnupdate = new DevExpress.XtraEditors.SimpleButton();
            this.btnAdd = new DevExpress.XtraEditors.SimpleButton();
            this.btnSelect = new DevExpress.XtraEditors.SimpleButton();
            this.btnDelete = new DevExpress.XtraEditors.SimpleButton();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem7 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem8 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem3 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlGroup2 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.label3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.label1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.lblid = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDefect)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem7)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem8)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.label3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.label1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            this.SuspendLayout();
            // 
            // txtid
            // 
            this.txtid.Location = new System.Drawing.Point(546, 37);
            this.txtid.Name = "txtid";
            this.txtid.Size = new System.Drawing.Size(244, 20);
            this.txtid.TabIndex = 3;
            // 
            // cbTestType
            // 
            this.cbTestType.FormattingEnabled = true;
            this.cbTestType.Location = new System.Drawing.Point(75, 37);
            this.cbTestType.Name = "cbTestType";
            this.cbTestType.Size = new System.Drawing.Size(204, 22);
            this.cbTestType.TabIndex = 0;
            this.cbTestType.SelectedIndexChanged += new System.EventHandler(this.cbTestType_SelectedIndexChanged);
            // 
            // txtTestItem
            // 
            this.txtTestItem.Location = new System.Drawing.Point(334, 37);
            this.txtTestItem.Name = "txtTestItem";
            this.txtTestItem.Size = new System.Drawing.Size(157, 20);
            this.txtTestItem.TabIndex = 2;
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(149, 74);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(121, 22);
            this.btnReset.StyleController = this.layoutControl1;
            this.btnReset.TabIndex = 4;
            this.btnReset.Text = "重 置";
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // layoutControl1
            // 
            this.layoutControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.layoutControl1.Controls.Add(this.txtid);
            this.layoutControl1.Controls.Add(this.dgvDefect);
            this.layoutControl1.Controls.Add(this.btnupdate);
            this.layoutControl1.Controls.Add(this.btnAdd);
            this.layoutControl1.Controls.Add(this.btnSelect);
            this.layoutControl1.Controls.Add(this.btnDelete);
            this.layoutControl1.Controls.Add(this.txtTestItem);
            this.layoutControl1.Controls.Add(this.btnReset);
            this.layoutControl1.Controls.Add(this.cbTestType);
            this.layoutControl1.Location = new System.Drawing.Point(53, 18);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.layoutControlGroup1;
            this.layoutControl1.Size = new System.Drawing.Size(814, 628);
            this.layoutControl1.TabIndex = 33;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // dgvDefect
            // 
            this.dgvDefect.Location = new System.Drawing.Point(12, 100);
            this.dgvDefect.MainView = this.gridView;
            this.dgvDefect.Name = "dgvDefect";
            this.dgvDefect.Size = new System.Drawing.Size(790, 516);
            this.dgvDefect.TabIndex = 9;
            this.dgvDefect.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView});
            // 
            // gridView
            // 
            this.gridView.Appearance.HeaderPanel.Options.UseTextOptions = true;
            this.gridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridView.Appearance.Row.Options.UseTextOptions = true;
            this.gridView.Appearance.Row.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridView.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.TestType,
            this.TestItem,
            this.Item});
            this.gridView.GridControl = this.dgvDefect;
            this.gridView.Name = "gridView";
            this.gridView.OptionsBehavior.Editable = false;
            this.gridView.OptionsView.ShowGroupPanel = false;
            this.gridView.Click += new System.EventHandler(this.gridView_Click);
            // 
            // TestType
            // 
            this.TestType.Caption = "测试类别";
            this.TestType.FieldName = "TestType";
            this.TestType.Name = "TestType";
            this.TestType.Visible = true;
            this.TestType.VisibleIndex = 0;
            // 
            // TestItem
            // 
            this.TestItem.Caption = "测试项目";
            this.TestItem.FieldName = "TestItem";
            this.TestItem.Name = "TestItem";
            this.TestItem.Visible = true;
            this.TestItem.VisibleIndex = 1;
            // 
            // Item
            // 
            this.Item.Caption = "顺序";
            this.Item.FieldName = "Item";
            this.Item.Name = "Item";
            this.Item.Visible = true;
            this.Item.VisibleIndex = 2;
            // 
            // btnupdate
            // 
            this.btnupdate.Location = new System.Drawing.Point(624, 74);
            this.btnupdate.Name = "btnupdate";
            this.btnupdate.Size = new System.Drawing.Size(125, 22);
            this.btnupdate.StyleController = this.layoutControl1;
            this.btnupdate.TabIndex = 8;
            this.btnupdate.Text = "更 新";
            this.btnupdate.Click += new System.EventHandler(this.btnupdate_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(396, 74);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(106, 22);
            this.btnAdd.StyleController = this.layoutControl1;
            this.btnAdd.TabIndex = 6;
            this.btnAdd.Text = "新 增";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnSelect
            // 
            this.btnSelect.Location = new System.Drawing.Point(274, 74);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(118, 22);
            this.btnSelect.StyleController = this.layoutControl1;
            this.btnSelect.TabIndex = 5;
            this.btnSelect.Text = "查 询";
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(506, 74);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(114, 22);
            this.btnDelete.StyleController = this.layoutControl1;
            this.btnDelete.TabIndex = 7;
            this.btnDelete.Text = "删 除";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.layoutControlGroup1.GroupBordersVisible = false;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem4,
            this.layoutControlItem7,
            this.layoutControlItem8,
            this.layoutControlItem5,
            this.layoutControlItem6,
            this.emptySpaceItem2,
            this.emptySpaceItem3,
            this.layoutControlGroup2,
            this.layoutControlItem1});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(814, 628);
            this.layoutControlGroup1.TextVisible = false;
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.btnReset;
            this.layoutControlItem4.Location = new System.Drawing.Point(137, 62);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(125, 26);
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // layoutControlItem7
            // 
            this.layoutControlItem7.Control = this.btnDelete;
            this.layoutControlItem7.Location = new System.Drawing.Point(494, 62);
            this.layoutControlItem7.Name = "layoutControlItem7";
            this.layoutControlItem7.Size = new System.Drawing.Size(118, 26);
            this.layoutControlItem7.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem7.TextVisible = false;
            // 
            // layoutControlItem8
            // 
            this.layoutControlItem8.Control = this.btnupdate;
            this.layoutControlItem8.Location = new System.Drawing.Point(612, 62);
            this.layoutControlItem8.Name = "layoutControlItem8";
            this.layoutControlItem8.Size = new System.Drawing.Size(129, 26);
            this.layoutControlItem8.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem8.TextVisible = false;
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.btnSelect;
            this.layoutControlItem5.Location = new System.Drawing.Point(262, 62);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(122, 26);
            this.layoutControlItem5.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem5.TextVisible = false;
            // 
            // layoutControlItem6
            // 
            this.layoutControlItem6.Control = this.btnAdd;
            this.layoutControlItem6.Location = new System.Drawing.Point(384, 62);
            this.layoutControlItem6.Name = "layoutControlItem6";
            this.layoutControlItem6.Size = new System.Drawing.Size(110, 26);
            this.layoutControlItem6.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem6.TextVisible = false;
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(0, 62);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(137, 26);
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem3
            // 
            this.emptySpaceItem3.AllowHotTrack = false;
            this.emptySpaceItem3.Location = new System.Drawing.Point(741, 62);
            this.emptySpaceItem3.Name = "emptySpaceItem3";
            this.emptySpaceItem3.Size = new System.Drawing.Size(53, 26);
            this.emptySpaceItem3.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlGroup2
            // 
            this.layoutControlGroup2.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.label3,
            this.label1,
            this.lblid});
            this.layoutControlGroup2.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup2.Name = "layoutControlGroup2";
            this.layoutControlGroup2.Size = new System.Drawing.Size(794, 62);
            this.layoutControlGroup2.Text = "查询条件";
            // 
            // label3
            // 
            this.label3.AppearanceItemCaption.Options.UseTextOptions = true;
            this.label3.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.label3.Control = this.cbTestType;
            this.label3.Location = new System.Drawing.Point(0, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(259, 25);
            this.label3.Text = "测试类别";
            this.label3.TextSize = new System.Drawing.Size(48, 14);
            // 
            // label1
            // 
            this.label1.AppearanceItemCaption.Options.UseTextOptions = true;
            this.label1.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.label1.Control = this.txtTestItem;
            this.label1.Location = new System.Drawing.Point(259, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(212, 25);
            this.label1.Text = "测试项目";
            this.label1.TextSize = new System.Drawing.Size(48, 14);
            // 
            // lblid
            // 
            this.lblid.AppearanceItemCaption.Options.UseTextOptions = true;
            this.lblid.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblid.Control = this.txtid;
            this.lblid.Location = new System.Drawing.Point(471, 0);
            this.lblid.Name = "lblid";
            this.lblid.Size = new System.Drawing.Size(299, 25);
            this.lblid.Text = "顺序";
            this.lblid.TextSize = new System.Drawing.Size(48, 14);
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.dgvDefect;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 88);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(794, 520);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // TestItemInfo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(925, 653);
            this.Controls.Add(this.layoutControl1);
            this.MaximizeBox = false;
            this.Name = "TestItemInfo";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "测试项目信息";
            this.Load += new System.EventHandler(this.TestItemInfo_Load);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvDefect)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem7)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem8)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.label3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.label1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.TextBox txtid;
        private System.Windows.Forms.ComboBox cbTestType;
        private System.Windows.Forms.TextBox txtTestItem;
        private DevExpress.XtraEditors.SimpleButton btnReset;
        private DevExpress.XtraEditors.SimpleButton btnSelect;
        private DevExpress.XtraEditors.SimpleButton btnAdd;
        private DevExpress.XtraEditors.SimpleButton btnDelete;
        private DevExpress.XtraEditors.SimpleButton btnupdate;
        private DevExpress.XtraGrid.GridControl dgvDefect;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView;
        private DevExpress.XtraGrid.Columns.GridColumn TestType;
        private DevExpress.XtraGrid.Columns.GridColumn TestItem;
        private DevExpress.XtraGrid.Columns.GridColumn Item;
        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem7;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem8;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem3;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup2;
        private DevExpress.XtraLayout.LayoutControlItem label3;
        private DevExpress.XtraLayout.LayoutControlItem label1;
        private DevExpress.XtraLayout.LayoutControlItem lblid;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
    }
}