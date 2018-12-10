namespace DX_QMS
{
    partial class ReliabilitySet
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
            this.txtleadtime = new System.Windows.Forms.NumericUpDown();
            this.txtcheckcycle = new System.Windows.Forms.NumericUpDown();
            this.lblinfo = new System.Windows.Forms.Label();
            this.cbTestType = new System.Windows.Forms.ComboBox();
            this.txtproductcode = new System.Windows.Forms.TextBox();
            this.btnsave = new DevExpress.XtraEditors.SimpleButton();
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.databind = new DevExpress.XtraGrid.GridControl();
            this.gridView = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.btndel = new DevExpress.XtraEditors.SimpleButton();
            this.btnReliabilitytypeAdd = new DevExpress.XtraEditors.SimpleButton();
            this.btnsearch = new DevExpress.XtraEditors.SimpleButton();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.lblproductcode = new DevExpress.XtraLayout.LayoutControlItem();
            this.label3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.lblcheckcycle = new DevExpress.XtraLayout.LayoutControlItem();
            this.lblleadtime = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem3 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.txtleadtime)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtcheckcycle)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.databind)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblproductcode)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.label3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblcheckcycle)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblleadtime)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            this.SuspendLayout();
            // 
            // txtleadtime
            // 
            this.txtleadtime.Location = new System.Drawing.Point(828, 12);
            this.txtleadtime.Maximum = new decimal(new int[] {
            365,
            0,
            0,
            0});
            this.txtleadtime.Name = "txtleadtime";
            this.txtleadtime.Size = new System.Drawing.Size(124, 22);
            this.txtleadtime.TabIndex = 33;
            this.txtleadtime.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // txtcheckcycle
            // 
            this.txtcheckcycle.Location = new System.Drawing.Point(619, 12);
            this.txtcheckcycle.Maximum = new decimal(new int[] {
            365,
            0,
            0,
            0});
            this.txtcheckcycle.Name = "txtcheckcycle";
            this.txtcheckcycle.Size = new System.Drawing.Size(132, 22);
            this.txtcheckcycle.TabIndex = 34;
            this.txtcheckcycle.Value = new decimal(new int[] {
            90,
            0,
            0,
            0});
            // 
            // lblinfo
            // 
            this.lblinfo.Location = new System.Drawing.Point(12, 64);
            this.lblinfo.Name = "lblinfo";
            this.lblinfo.Size = new System.Drawing.Size(940, 20);
            this.lblinfo.TabIndex = 30;
            // 
            // cbTestType
            // 
            this.cbTestType.FormattingEnabled = true;
            this.cbTestType.Location = new System.Drawing.Point(322, 12);
            this.cbTestType.Name = "cbTestType";
            this.cbTestType.Size = new System.Drawing.Size(161, 22);
            this.cbTestType.TabIndex = 28;
            // 
            // txtproductcode
            // 
            this.txtproductcode.Location = new System.Drawing.Point(85, 12);
            this.txtproductcode.Name = "txtproductcode";
            this.txtproductcode.Size = new System.Drawing.Size(160, 20);
            this.txtproductcode.TabIndex = 26;
            this.txtproductcode.KeyUp += new System.Windows.Forms.KeyEventHandler(this.txtproductcode_KeyUp);
            // 
            // btnsave
            // 
            this.btnsave.Location = new System.Drawing.Point(265, 38);
            this.btnsave.Name = "btnsave";
            this.btnsave.Size = new System.Drawing.Size(162, 22);
            this.btnsave.StyleController = this.layoutControl1;
            this.btnsave.TabIndex = 44;
            this.btnsave.Text = "保 存";
            this.btnsave.Click += new System.EventHandler(this.btnsave_Click);
            // 
            // layoutControl1
            // 
            this.layoutControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.layoutControl1.Controls.Add(this.databind);
            this.layoutControl1.Controls.Add(this.lblinfo);
            this.layoutControl1.Controls.Add(this.btndel);
            this.layoutControl1.Controls.Add(this.btnReliabilitytypeAdd);
            this.layoutControl1.Controls.Add(this.btnsearch);
            this.layoutControl1.Controls.Add(this.btnsave);
            this.layoutControl1.Controls.Add(this.txtleadtime);
            this.layoutControl1.Controls.Add(this.txtproductcode);
            this.layoutControl1.Controls.Add(this.cbTestType);
            this.layoutControl1.Controls.Add(this.txtcheckcycle);
            this.layoutControl1.Location = new System.Drawing.Point(52, 20);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.layoutControlGroup1;
            this.layoutControl1.Size = new System.Drawing.Size(964, 591);
            this.layoutControl1.TabIndex = 49;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // databind
            // 
            this.databind.Location = new System.Drawing.Point(12, 88);
            this.databind.MainView = this.gridView;
            this.databind.Name = "databind";
            this.databind.Size = new System.Drawing.Size(940, 491);
            this.databind.TabIndex = 48;
            this.databind.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView});
            // 
            // gridView
            // 
            this.gridView.Appearance.HeaderPanel.Options.UseTextOptions = true;
            this.gridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridView.GridControl = this.databind;
            this.gridView.Name = "gridView";
            this.gridView.OptionsBehavior.Editable = false;
            this.gridView.OptionsSelection.MultiSelect = true;
            this.gridView.OptionsView.ShowGroupPanel = false;
            this.gridView.Click += new System.EventHandler(this.gridView_Click);
            // 
            // btndel
            // 
            this.btndel.Location = new System.Drawing.Point(586, 38);
            this.btndel.Name = "btndel";
            this.btndel.Size = new System.Drawing.Size(135, 22);
            this.btndel.StyleController = this.layoutControl1;
            this.btndel.TabIndex = 46;
            this.btndel.Text = "删 除";
            this.btndel.Click += new System.EventHandler(this.btndel_Click);
            // 
            // btnReliabilitytypeAdd
            // 
            this.btnReliabilitytypeAdd.Location = new System.Drawing.Point(487, 12);
            this.btnReliabilitytypeAdd.Name = "btnReliabilitytypeAdd";
            this.btnReliabilitytypeAdd.Size = new System.Drawing.Size(55, 22);
            this.btnReliabilitytypeAdd.StyleController = this.layoutControl1;
            this.btnReliabilitytypeAdd.TabIndex = 47;
            this.btnReliabilitytypeAdd.Text = "...";
            this.btnReliabilitytypeAdd.Click += new System.EventHandler(this.btnReliabilitytypeAdd_Click);
            // 
            // btnsearch
            // 
            this.btnsearch.Location = new System.Drawing.Point(431, 38);
            this.btnsearch.Name = "btnsearch";
            this.btnsearch.Size = new System.Drawing.Size(151, 22);
            this.btnsearch.StyleController = this.layoutControl1;
            this.btnsearch.TabIndex = 45;
            this.btnsearch.Text = "查 询";
            this.btnsearch.Click += new System.EventHandler(this.btnsearch_Click);
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.layoutControlGroup1.GroupBordersVisible = false;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.lblproductcode,
            this.label3,
            this.layoutControlItem3,
            this.lblcheckcycle,
            this.lblleadtime,
            this.layoutControlItem1,
            this.layoutControlItem2,
            this.layoutControlItem4,
            this.emptySpaceItem2,
            this.emptySpaceItem3,
            this.layoutControlItem5,
            this.layoutControlItem6});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(964, 591);
            this.layoutControlGroup1.TextVisible = false;
            // 
            // lblproductcode
            // 
            this.lblproductcode.AppearanceItemCaption.Options.UseTextOptions = true;
            this.lblproductcode.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblproductcode.Control = this.txtproductcode;
            this.lblproductcode.Location = new System.Drawing.Point(0, 0);
            this.lblproductcode.Name = "lblproductcode";
            this.lblproductcode.Size = new System.Drawing.Size(237, 26);
            this.lblproductcode.Text = "物料编码";
            this.lblproductcode.TextSize = new System.Drawing.Size(70, 14);
            // 
            // label3
            // 
            this.label3.AppearanceItemCaption.Options.UseTextOptions = true;
            this.label3.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.label3.Control = this.cbTestType;
            this.label3.Location = new System.Drawing.Point(237, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(238, 26);
            this.label3.Text = "测试类别";
            this.label3.TextSize = new System.Drawing.Size(70, 14);
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.btnReliabilitytypeAdd;
            this.layoutControlItem3.Location = new System.Drawing.Point(475, 0);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(59, 26);
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // lblcheckcycle
            // 
            this.lblcheckcycle.AppearanceItemCaption.Options.UseTextOptions = true;
            this.lblcheckcycle.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblcheckcycle.Control = this.txtcheckcycle;
            this.lblcheckcycle.Location = new System.Drawing.Point(534, 0);
            this.lblcheckcycle.Name = "lblcheckcycle";
            this.lblcheckcycle.Size = new System.Drawing.Size(209, 26);
            this.lblcheckcycle.Text = "测试周期(天)";
            this.lblcheckcycle.TextSize = new System.Drawing.Size(70, 14);
            // 
            // lblleadtime
            // 
            this.lblleadtime.AppearanceItemCaption.Options.UseTextOptions = true;
            this.lblleadtime.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblleadtime.Control = this.txtleadtime;
            this.lblleadtime.Location = new System.Drawing.Point(743, 0);
            this.lblleadtime.Name = "lblleadtime";
            this.lblleadtime.Size = new System.Drawing.Size(201, 26);
            this.lblleadtime.Text = "提前期(天)";
            this.lblleadtime.TextSize = new System.Drawing.Size(70, 14);
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.btnsave;
            this.layoutControlItem1.Location = new System.Drawing.Point(253, 26);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(166, 26);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.btnsearch;
            this.layoutControlItem2.Location = new System.Drawing.Point(419, 26);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(155, 26);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.btndel;
            this.layoutControlItem4.Location = new System.Drawing.Point(574, 26);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(139, 26);
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(0, 26);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(253, 26);
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem3
            // 
            this.emptySpaceItem3.AllowHotTrack = false;
            this.emptySpaceItem3.Location = new System.Drawing.Point(713, 26);
            this.emptySpaceItem3.Name = "emptySpaceItem3";
            this.emptySpaceItem3.Size = new System.Drawing.Size(231, 26);
            this.emptySpaceItem3.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.lblinfo;
            this.layoutControlItem5.Location = new System.Drawing.Point(0, 52);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(944, 24);
            this.layoutControlItem5.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem5.TextVisible = false;
            // 
            // layoutControlItem6
            // 
            this.layoutControlItem6.Control = this.databind;
            this.layoutControlItem6.Location = new System.Drawing.Point(0, 76);
            this.layoutControlItem6.Name = "layoutControlItem6";
            this.layoutControlItem6.Size = new System.Drawing.Size(944, 495);
            this.layoutControlItem6.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem6.TextVisible = false;
            // 
            // ReliabilitySet
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1074, 623);
            this.Controls.Add(this.layoutControl1);
            this.MaximizeBox = false;
            this.Name = "ReliabilitySet";
            this.Text = "可靠性/真实性测试设置";
            ((System.ComponentModel.ISupportInitialize)(this.txtleadtime)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtcheckcycle)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.databind)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblproductcode)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.label3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblcheckcycle)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblleadtime)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.NumericUpDown txtleadtime;
        private System.Windows.Forms.NumericUpDown txtcheckcycle;
        private System.Windows.Forms.Label lblinfo;
        private System.Windows.Forms.ComboBox cbTestType;
        private System.Windows.Forms.TextBox txtproductcode;
        private DevExpress.XtraEditors.SimpleButton btnsave;
        private DevExpress.XtraEditors.SimpleButton btnsearch;
        private DevExpress.XtraEditors.SimpleButton btndel;
        private DevExpress.XtraEditors.SimpleButton btnReliabilitytypeAdd;
        private DevExpress.XtraGrid.GridControl databind;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView;
        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem lblproductcode;
        private DevExpress.XtraLayout.LayoutControlItem label3;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraLayout.LayoutControlItem lblcheckcycle;
        private DevExpress.XtraLayout.LayoutControlItem lblleadtime;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem3;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
    }
}