namespace DX_QMS
{
    partial class OQCTestType
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
            this.code = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtid = new System.Windows.Forms.TextBox();
            this.cbTestType = new System.Windows.Forms.ComboBox();
            this.txtTestItem = new System.Windows.Forms.TextBox();
            this.Definevalue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Definetype = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgvDefect = new System.Windows.Forms.DataGridView();
            this.sort = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtconment = new System.Windows.Forms.TextBox();
            this.groupControl1 = new DevExpress.XtraEditors.GroupControl();
            this.btnupdate = new DevExpress.XtraEditors.SimpleButton();
            this.btnDelete = new DevExpress.XtraEditors.SimpleButton();
            this.btnAdd = new DevExpress.XtraEditors.SimpleButton();
            this.btnSelect = new DevExpress.XtraEditors.SimpleButton();
            this.btnReset = new DevExpress.XtraEditors.SimpleButton();
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControl2 = new DevExpress.XtraLayout.LayoutControl();
            this.layoutControlGroup2 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.label3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.label1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.label2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.lblid = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem7 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem3 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem4 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem5 = new DevExpress.XtraLayout.EmptySpaceItem();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDefect)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl1)).BeginInit();
            this.groupControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl2)).BeginInit();
            this.layoutControl2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.label3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.label1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.label2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem7)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem5)).BeginInit();
            this.SuspendLayout();
            // 
            // code
            // 
            this.code.DataPropertyName = "code";
            this.code.HeaderText = "内容";
            this.code.Name = "code";
            this.code.ReadOnly = true;
            // 
            // txtid
            // 
            this.txtid.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtid.Location = new System.Drawing.Point(361, 48);
            this.txtid.Name = "txtid";
            this.txtid.Size = new System.Drawing.Size(366, 20);
            this.txtid.TabIndex = 2;
            // 
            // cbTestType
            // 
            this.cbTestType.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.cbTestType.FormattingEnabled = true;
            this.cbTestType.Location = new System.Drawing.Point(64, 12);
            this.cbTestType.Name = "cbTestType";
            this.cbTestType.Size = new System.Drawing.Size(241, 22);
            this.cbTestType.TabIndex = 5;
            this.cbTestType.SelectedIndexChanged += new System.EventHandler(this.cbTestType_SelectedIndexChanged);
            // 
            // txtTestItem
            // 
            this.txtTestItem.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtTestItem.Location = new System.Drawing.Point(361, 12);
            this.txtTestItem.Name = "txtTestItem";
            this.txtTestItem.Size = new System.Drawing.Size(366, 20);
            this.txtTestItem.TabIndex = 2;
            // 
            // Definevalue
            // 
            this.Definevalue.DataPropertyName = "Definevalue";
            this.Definevalue.HeaderText = "类别名称";
            this.Definevalue.Name = "Definevalue";
            this.Definevalue.ReadOnly = true;
            // 
            // Definetype
            // 
            this.Definetype.DataPropertyName = "Definetype";
            this.Definetype.HeaderText = "类别";
            this.Definetype.Name = "Definetype";
            this.Definetype.ReadOnly = true;
            // 
            // dgvDefect
            // 
            this.dgvDefect.AllowUserToAddRows = false;
            this.dgvDefect.AllowUserToDeleteRows = false;
            this.dgvDefect.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.dgvDefect.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvDefect.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            this.dgvDefect.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDefect.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Definetype,
            this.Definevalue,
            this.code,
            this.sort});
            this.dgvDefect.Location = new System.Drawing.Point(12, 127);
            this.dgvDefect.MultiSelect = false;
            this.dgvDefect.Name = "dgvDefect";
            this.dgvDefect.ReadOnly = true;
            this.dgvDefect.RowHeadersVisible = false;
            this.dgvDefect.RowTemplate.Height = 23;
            this.dgvDefect.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvDefect.Size = new System.Drawing.Size(715, 516);
            this.dgvDefect.TabIndex = 20;
            this.dgvDefect.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvDefect_CellClick);
            // 
            // sort
            // 
            this.sort.DataPropertyName = "sort";
            this.sort.HeaderText = "顺序";
            this.sort.Name = "sort";
            this.sort.ReadOnly = true;
            // 
            // txtconment
            // 
            this.txtconment.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtconment.Location = new System.Drawing.Point(64, 48);
            this.txtconment.Multiline = true;
            this.txtconment.Name = "txtconment";
            this.txtconment.Size = new System.Drawing.Size(241, 20);
            this.txtconment.TabIndex = 2;
            // 
            // groupControl1
            // 
            this.groupControl1.Controls.Add(this.layoutControl2);
            this.groupControl1.Location = new System.Drawing.Point(12, 12);
            this.groupControl1.Name = "groupControl1";
            this.groupControl1.Size = new System.Drawing.Size(928, 711);
            this.groupControl1.TabIndex = 23;
            this.groupControl1.Text = "查询条件";
            // 
            // btnupdate
            // 
            this.btnupdate.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnupdate.Location = new System.Drawing.Point(472, 90);
            this.btnupdate.MaximumSize = new System.Drawing.Size(80, 35);
            this.btnupdate.Name = "btnupdate";
            this.btnupdate.Size = new System.Drawing.Size(80, 22);
            this.btnupdate.StyleController = this.layoutControl2;
            this.btnupdate.TabIndex = 28;
            this.btnupdate.Text = "更 新";
            this.btnupdate.Click += new System.EventHandler(this.btnupdate_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnDelete.Location = new System.Drawing.Point(388, 90);
            this.btnDelete.MaximumSize = new System.Drawing.Size(80, 35);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(80, 22);
            this.btnDelete.StyleController = this.layoutControl2;
            this.btnDelete.TabIndex = 27;
            this.btnDelete.Text = "删 除";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnAdd.Location = new System.Drawing.Point(304, 90);
            this.btnAdd.MaximumSize = new System.Drawing.Size(80, 35);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(80, 22);
            this.btnAdd.StyleController = this.layoutControl2;
            this.btnAdd.TabIndex = 26;
            this.btnAdd.Text = "新 增";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnSelect
            // 
            this.btnSelect.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnSelect.Location = new System.Drawing.Point(221, 90);
            this.btnSelect.MaximumSize = new System.Drawing.Size(80, 35);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(79, 22);
            this.btnSelect.StyleController = this.layoutControl2;
            this.btnSelect.TabIndex = 25;
            this.btnSelect.Text = "查 询";
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // btnReset
            // 
            this.btnReset.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnReset.Location = new System.Drawing.Point(139, 90);
            this.btnReset.MaximumSize = new System.Drawing.Size(80, 35);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(78, 22);
            this.btnReset.StyleController = this.layoutControl2;
            this.btnReset.TabIndex = 24;
            this.btnReset.Text = "重 置";
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.groupControl1);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.layoutControlGroup1;
            this.layoutControl1.Size = new System.Drawing.Size(952, 735);
            this.layoutControl1.TabIndex = 26;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.layoutControlGroup1.GroupBordersVisible = false;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(952, 735);
            this.layoutControlGroup1.TextVisible = false;
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.groupControl1;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(932, 715);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // layoutControl2
            // 
            this.layoutControl2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.layoutControl2.Controls.Add(this.btnupdate);
            this.layoutControl2.Controls.Add(this.btnDelete);
            this.layoutControl2.Controls.Add(this.dgvDefect);
            this.layoutControl2.Controls.Add(this.btnAdd);
            this.layoutControl2.Controls.Add(this.btnSelect);
            this.layoutControl2.Controls.Add(this.txtid);
            this.layoutControl2.Controls.Add(this.btnReset);
            this.layoutControl2.Controls.Add(this.cbTestType);
            this.layoutControl2.Controls.Add(this.txtTestItem);
            this.layoutControl2.Controls.Add(this.txtconment);
            this.layoutControl2.Location = new System.Drawing.Point(77, 43);
            this.layoutControl2.Name = "layoutControl2";
            this.layoutControl2.Root = this.layoutControlGroup2;
            this.layoutControl2.Size = new System.Drawing.Size(739, 655);
            this.layoutControl2.TabIndex = 29;
            this.layoutControl2.Text = "layoutControl2";
            // 
            // layoutControlGroup2
            // 
            this.layoutControlGroup2.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.layoutControlGroup2.GroupBordersVisible = false;
            this.layoutControlGroup2.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.label3,
            this.emptySpaceItem1,
            this.label1,
            this.label2,
            this.lblid,
            this.layoutControlItem3,
            this.layoutControlItem4,
            this.layoutControlItem5,
            this.layoutControlItem6,
            this.layoutControlItem7,
            this.emptySpaceItem2,
            this.emptySpaceItem3,
            this.layoutControlItem2,
            this.emptySpaceItem4,
            this.emptySpaceItem5});
            this.layoutControlGroup2.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup2.Name = "layoutControlGroup2";
            this.layoutControlGroup2.Size = new System.Drawing.Size(739, 655);
            this.layoutControlGroup2.TextVisible = false;
            // 
            // label3
            // 
            this.label3.AppearanceItemCaption.Options.UseTextOptions = true;
            this.label3.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.label3.Control = this.cbTestType;
            this.label3.Location = new System.Drawing.Point(0, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(297, 25);
            this.label3.Text = "类别";
            this.label3.TextSize = new System.Drawing.Size(48, 14);
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(0, 104);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(719, 11);
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // label1
            // 
            this.label1.AppearanceItemCaption.Options.UseTextOptions = true;
            this.label1.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.label1.Control = this.txtTestItem;
            this.label1.Location = new System.Drawing.Point(297, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(422, 25);
            this.label1.Text = "类别名称";
            this.label1.TextSize = new System.Drawing.Size(48, 14);
            // 
            // label2
            // 
            this.label2.AppearanceItemCaption.Options.UseTextOptions = true;
            this.label2.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.label2.Control = this.txtconment;
            this.label2.Location = new System.Drawing.Point(0, 36);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(297, 24);
            this.label2.Text = "内容";
            this.label2.TextSize = new System.Drawing.Size(48, 14);
            // 
            // lblid
            // 
            this.lblid.AppearanceItemCaption.Options.UseTextOptions = true;
            this.lblid.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblid.Control = this.txtid;
            this.lblid.Location = new System.Drawing.Point(297, 36);
            this.lblid.Name = "lblid";
            this.lblid.Size = new System.Drawing.Size(422, 24);
            this.lblid.Text = "顺序";
            this.lblid.TextSize = new System.Drawing.Size(48, 14);
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.btnReset;
            this.layoutControlItem3.Location = new System.Drawing.Point(127, 78);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(82, 26);
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.btnSelect;
            this.layoutControlItem4.Location = new System.Drawing.Point(209, 78);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(83, 26);
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.btnAdd;
            this.layoutControlItem5.Location = new System.Drawing.Point(292, 78);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(84, 26);
            this.layoutControlItem5.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem5.TextVisible = false;
            // 
            // layoutControlItem6
            // 
            this.layoutControlItem6.Control = this.btnDelete;
            this.layoutControlItem6.Location = new System.Drawing.Point(376, 78);
            this.layoutControlItem6.Name = "layoutControlItem6";
            this.layoutControlItem6.Size = new System.Drawing.Size(84, 26);
            this.layoutControlItem6.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem6.TextVisible = false;
            // 
            // layoutControlItem7
            // 
            this.layoutControlItem7.Control = this.btnupdate;
            this.layoutControlItem7.Location = new System.Drawing.Point(460, 78);
            this.layoutControlItem7.Name = "layoutControlItem7";
            this.layoutControlItem7.Size = new System.Drawing.Size(84, 26);
            this.layoutControlItem7.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem7.TextVisible = false;
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(0, 78);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(127, 26);
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem3
            // 
            this.emptySpaceItem3.AllowHotTrack = false;
            this.emptySpaceItem3.Location = new System.Drawing.Point(544, 78);
            this.emptySpaceItem3.Name = "emptySpaceItem3";
            this.emptySpaceItem3.Size = new System.Drawing.Size(175, 26);
            this.emptySpaceItem3.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.dgvDefect;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 115);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(719, 520);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // emptySpaceItem4
            // 
            this.emptySpaceItem4.AllowHotTrack = false;
            this.emptySpaceItem4.Location = new System.Drawing.Point(0, 25);
            this.emptySpaceItem4.Name = "emptySpaceItem4";
            this.emptySpaceItem4.Size = new System.Drawing.Size(719, 11);
            this.emptySpaceItem4.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem5
            // 
            this.emptySpaceItem5.AllowHotTrack = false;
            this.emptySpaceItem5.Location = new System.Drawing.Point(0, 60);
            this.emptySpaceItem5.Name = "emptySpaceItem5";
            this.emptySpaceItem5.Size = new System.Drawing.Size(719, 18);
            this.emptySpaceItem5.TextSize = new System.Drawing.Size(0, 0);
            // 
            // OQCTestType
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(952, 735);
            this.Controls.Add(this.layoutControl1);
            this.Name = "OQCTestType";
            this.Text = "作业设置";
            this.Load += new System.EventHandler(this.OQCTestType_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvDefect)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl1)).EndInit();
            this.groupControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl2)).EndInit();
            this.layoutControl2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.label3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.label1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.label2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem7)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem5)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.DataGridViewTextBoxColumn code;
        private System.Windows.Forms.TextBox txtid;
        private System.Windows.Forms.ComboBox cbTestType;
        private System.Windows.Forms.TextBox txtTestItem;
        private System.Windows.Forms.DataGridViewTextBoxColumn Definevalue;
        private System.Windows.Forms.DataGridViewTextBoxColumn Definetype;
        private System.Windows.Forms.DataGridView dgvDefect;
        private System.Windows.Forms.DataGridViewTextBoxColumn sort;
        private System.Windows.Forms.TextBox txtconment;
        private DevExpress.XtraEditors.GroupControl groupControl1;
        private DevExpress.XtraEditors.SimpleButton btnupdate;
        private DevExpress.XtraEditors.SimpleButton btnDelete;
        private DevExpress.XtraEditors.SimpleButton btnAdd;
        private DevExpress.XtraEditors.SimpleButton btnSelect;
        private DevExpress.XtraEditors.SimpleButton btnReset;
        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControl layoutControl2;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup2;
        private DevExpress.XtraLayout.LayoutControlItem label3;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraLayout.LayoutControlItem label1;
        private DevExpress.XtraLayout.LayoutControlItem label2;
        private DevExpress.XtraLayout.LayoutControlItem lblid;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem7;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem3;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem4;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem5;
    }
}