namespace DX_QMS
{
    partial class TestBigTypeSet
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
            this.txtTestType = new System.Windows.Forms.TextBox();
            this.databind_father = new DevExpress.XtraGrid.GridControl();
            this.databind = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.btnquery = new DevExpress.XtraEditors.SimpleButton();
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.btnupdate = new DevExpress.XtraEditors.SimpleButton();
            this.btndel = new DevExpress.XtraEditors.SimpleButton();
            this.btnSave = new DevExpress.XtraEditors.SimpleButton();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.gbDefect = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem3 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.label1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.databind_father)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.databind)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gbDefect)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.label1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            this.SuspendLayout();
            // 
            // txtTestType
            // 
            this.txtTestType.Location = new System.Drawing.Point(75, 37);
            this.txtTestType.Name = "txtTestType";
            this.txtTestType.Size = new System.Drawing.Size(292, 20);
            this.txtTestType.TabIndex = 6;
            // 
            // databind_father
            // 
            this.databind_father.Location = new System.Drawing.Point(12, 99);
            this.databind_father.MainView = this.databind;
            this.databind_father.Name = "databind_father";
            this.databind_father.ShowOnlyPredefinedDetails = true;
            this.databind_father.Size = new System.Drawing.Size(367, 420);
            this.databind_father.TabIndex = 8;
            this.databind_father.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.databind});
            this.databind_father.Click += new System.EventHandler(this.databind_father_Click);
            // 
            // databind
            // 
            this.databind.GridControl = this.databind_father;
            this.databind.Name = "databind";
            this.databind.OptionsBehavior.Editable = false;
            this.databind.OptionsView.ShowGroupPanel = false;
            // 
            // btnquery
            // 
            this.btnquery.Location = new System.Drawing.Point(272, 61);
            this.btnquery.Name = "btnquery";
            this.btnquery.Size = new System.Drawing.Size(58, 22);
            this.btnquery.StyleController = this.layoutControl1;
            this.btnquery.TabIndex = 24;
            this.btnquery.Text = "查 询";
            this.btnquery.Click += new System.EventHandler(this.btnquery_Click);
            // 
            // layoutControl1
            // 
            this.layoutControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.layoutControl1.Controls.Add(this.btnquery);
            this.layoutControl1.Controls.Add(this.databind_father);
            this.layoutControl1.Controls.Add(this.btnupdate);
            this.layoutControl1.Controls.Add(this.btndel);
            this.layoutControl1.Controls.Add(this.btnSave);
            this.layoutControl1.Controls.Add(this.txtTestType);
            this.layoutControl1.Location = new System.Drawing.Point(57, 22);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.layoutControlGroup1;
            this.layoutControl1.Size = new System.Drawing.Size(391, 531);
            this.layoutControl1.TabIndex = 21;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // btnupdate
            // 
            this.btnupdate.Location = new System.Drawing.Point(193, 61);
            this.btnupdate.Name = "btnupdate";
            this.btnupdate.Size = new System.Drawing.Size(75, 22);
            this.btnupdate.StyleController = this.layoutControl1;
            this.btnupdate.TabIndex = 23;
            this.btnupdate.Text = "更 新";
            this.btnupdate.Click += new System.EventHandler(this.btnupdate_Click);
            // 
            // btndel
            // 
            this.btndel.Location = new System.Drawing.Point(126, 61);
            this.btndel.Name = "btndel";
            this.btndel.Size = new System.Drawing.Size(63, 22);
            this.btndel.StyleController = this.layoutControl1;
            this.btndel.TabIndex = 22;
            this.btndel.Text = "删 除";
            this.btndel.Click += new System.EventHandler(this.btndel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(67, 61);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(55, 22);
            this.btnSave.StyleController = this.layoutControl1;
            this.btnSave.TabIndex = 21;
            this.btnSave.Text = "保 存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.layoutControlGroup1.GroupBordersVisible = false;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.gbDefect,
            this.layoutControlItem5});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(391, 531);
            this.layoutControlGroup1.TextVisible = false;
            // 
            // gbDefect
            // 
            this.gbDefect.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.layoutControlItem2,
            this.layoutControlItem3,
            this.layoutControlItem4,
            this.emptySpaceItem3,
            this.emptySpaceItem2,
            this.label1});
            this.gbDefect.Location = new System.Drawing.Point(0, 0);
            this.gbDefect.Name = "gbDefect";
            this.gbDefect.Size = new System.Drawing.Size(371, 87);
            this.gbDefect.Text = "测试类别维护";
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.btnSave;
            this.layoutControlItem1.Location = new System.Drawing.Point(43, 24);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(59, 26);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.btndel;
            this.layoutControlItem2.Location = new System.Drawing.Point(102, 24);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(67, 26);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.btnupdate;
            this.layoutControlItem3.Location = new System.Drawing.Point(169, 24);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(79, 26);
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.btnquery;
            this.layoutControlItem4.Location = new System.Drawing.Point(248, 24);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(62, 26);
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // emptySpaceItem3
            // 
            this.emptySpaceItem3.AllowHotTrack = false;
            this.emptySpaceItem3.Location = new System.Drawing.Point(310, 24);
            this.emptySpaceItem3.Name = "emptySpaceItem3";
            this.emptySpaceItem3.Size = new System.Drawing.Size(37, 26);
            this.emptySpaceItem3.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(0, 24);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(43, 26);
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // label1
            // 
            this.label1.Control = this.txtTestType;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(347, 24);
            this.label1.Text = "测试类别";
            this.label1.TextSize = new System.Drawing.Size(48, 14);
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.databind_father;
            this.layoutControlItem5.Location = new System.Drawing.Point(0, 87);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(371, 424);
            this.layoutControlItem5.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem5.TextVisible = false;
            // 
            // TestBigTypeSet
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(507, 574);
            this.Controls.Add(this.layoutControl1);
            this.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "TestBigTypeSet";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "测试类别维护";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.TestBigTypeSet_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.databind_father)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.databind)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gbDefect)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.label1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.TextBox txtTestType;
        private DevExpress.XtraGrid.GridControl databind_father;
        private DevExpress.XtraGrid.Views.Grid.GridView databind;
        private DevExpress.XtraEditors.SimpleButton btnSave;
        private DevExpress.XtraEditors.SimpleButton btndel;
        private DevExpress.XtraEditors.SimpleButton btnupdate;
        private DevExpress.XtraEditors.SimpleButton btnquery;
        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlGroup gbDefect;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem3;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.LayoutControlItem label1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
    }
}