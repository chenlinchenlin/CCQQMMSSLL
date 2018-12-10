namespace DX_QMS
{
    partial class TestSampleList
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
            this.txtqty = new System.Windows.Forms.TextBox();
            this.txtsupp = new System.Windows.Forms.TextBox();
            this.txtproductcode = new System.Windows.Forms.TextBox();
            this.txtremarks = new System.Windows.Forms.TextBox();
            this.txtsampletype = new System.Windows.Forms.TextBox();
            this.databind = new DevExpress.XtraGrid.GridControl();
            this.gridView = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.btnAdd = new DevExpress.XtraEditors.SimpleButton();
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.btnDel = new DevExpress.XtraEditors.SimpleButton();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.lblsampletype = new DevExpress.XtraLayout.LayoutControlItem();
            this.lblproductcode = new DevExpress.XtraLayout.LayoutControlItem();
            this.lblqty = new DevExpress.XtraLayout.LayoutControlItem();
            this.lblremarks = new DevExpress.XtraLayout.LayoutControlItem();
            this.label1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.databind)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblsampletype)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblproductcode)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblqty)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblremarks)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.label1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            this.SuspendLayout();
            // 
            // txtqty
            // 
            this.txtqty.Location = new System.Drawing.Point(598, 12);
            this.txtqty.Name = "txtqty";
            this.txtqty.Size = new System.Drawing.Size(246, 20);
            this.txtqty.TabIndex = 9;
            // 
            // txtsupp
            // 
            this.txtsupp.Location = new System.Drawing.Point(598, 36);
            this.txtsupp.Name = "txtsupp";
            this.txtsupp.ReadOnly = true;
            this.txtsupp.Size = new System.Drawing.Size(246, 20);
            this.txtsupp.TabIndex = 10;
            // 
            // txtproductcode
            // 
            this.txtproductcode.Location = new System.Drawing.Point(337, 12);
            this.txtproductcode.Name = "txtproductcode";
            this.txtproductcode.ReadOnly = true;
            this.txtproductcode.Size = new System.Drawing.Size(206, 20);
            this.txtproductcode.TabIndex = 11;
            // 
            // txtremarks
            // 
            this.txtremarks.Location = new System.Drawing.Point(63, 36);
            this.txtremarks.Name = "txtremarks";
            this.txtremarks.Size = new System.Drawing.Size(480, 20);
            this.txtremarks.TabIndex = 12;
            // 
            // txtsampletype
            // 
            this.txtsampletype.Location = new System.Drawing.Point(63, 12);
            this.txtsampletype.Name = "txtsampletype";
            this.txtsampletype.ReadOnly = true;
            this.txtsampletype.Size = new System.Drawing.Size(219, 20);
            this.txtsampletype.TabIndex = 13;
            // 
            // databind
            // 
            this.databind.Location = new System.Drawing.Point(12, 86);
            this.databind.MainView = this.gridView;
            this.databind.Name = "databind";
            this.databind.Size = new System.Drawing.Size(832, 611);
            this.databind.TabIndex = 19;
            this.databind.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView});
            // 
            // gridView
            // 
            this.gridView.Appearance.HeaderPanel.Options.UseTextOptions = true;
            this.gridView.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.gridView.GridControl = this.databind;
            this.gridView.Name = "gridView";
            this.gridView.OptionsView.ShowGroupPanel = false;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(115, 60);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(286, 22);
            this.btnAdd.StyleController = this.layoutControl1;
            this.btnAdd.TabIndex = 20;
            this.btnAdd.Text = "新增";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // layoutControl1
            // 
            this.layoutControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.layoutControl1.Controls.Add(this.btnAdd);
            this.layoutControl1.Controls.Add(this.databind);
            this.layoutControl1.Controls.Add(this.btnDel);
            this.layoutControl1.Controls.Add(this.txtsupp);
            this.layoutControl1.Controls.Add(this.txtqty);
            this.layoutControl1.Controls.Add(this.txtsampletype);
            this.layoutControl1.Controls.Add(this.txtremarks);
            this.layoutControl1.Controls.Add(this.txtproductcode);
            this.layoutControl1.Location = new System.Drawing.Point(85, 27);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.layoutControlGroup1;
            this.layoutControl1.Size = new System.Drawing.Size(856, 709);
            this.layoutControl1.TabIndex = 22;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(405, 60);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(265, 22);
            this.btnDel.StyleController = this.layoutControl1;
            this.btnDel.TabIndex = 21;
            this.btnDel.Text = "删除";
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.layoutControlGroup1.GroupBordersVisible = false;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.lblsampletype,
            this.lblproductcode,
            this.lblqty,
            this.lblremarks,
            this.label1,
            this.layoutControlItem2,
            this.layoutControlItem3,
            this.emptySpaceItem2,
            this.emptySpaceItem1,
            this.layoutControlItem1});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(856, 709);
            this.layoutControlGroup1.TextVisible = false;
            // 
            // lblsampletype
            // 
            this.lblsampletype.AppearanceItemCaption.Options.UseTextOptions = true;
            this.lblsampletype.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblsampletype.Control = this.txtsampletype;
            this.lblsampletype.Location = new System.Drawing.Point(0, 0);
            this.lblsampletype.Name = "lblsampletype";
            this.lblsampletype.Size = new System.Drawing.Size(274, 24);
            this.lblsampletype.Text = "样品类别";
            this.lblsampletype.TextSize = new System.Drawing.Size(48, 14);
            // 
            // lblproductcode
            // 
            this.lblproductcode.AppearanceItemCaption.Options.UseTextOptions = true;
            this.lblproductcode.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblproductcode.Control = this.txtproductcode;
            this.lblproductcode.Location = new System.Drawing.Point(274, 0);
            this.lblproductcode.Name = "lblproductcode";
            this.lblproductcode.Size = new System.Drawing.Size(261, 24);
            this.lblproductcode.Text = "编码";
            this.lblproductcode.TextSize = new System.Drawing.Size(48, 14);
            // 
            // lblqty
            // 
            this.lblqty.AppearanceItemCaption.Options.UseTextOptions = true;
            this.lblqty.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblqty.Control = this.txtqty;
            this.lblqty.Location = new System.Drawing.Point(535, 0);
            this.lblqty.Name = "lblqty";
            this.lblqty.Size = new System.Drawing.Size(301, 24);
            this.lblqty.Text = "使用数量";
            this.lblqty.TextSize = new System.Drawing.Size(48, 14);
            // 
            // lblremarks
            // 
            this.lblremarks.AppearanceItemCaption.Options.UseTextOptions = true;
            this.lblremarks.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblremarks.Control = this.txtremarks;
            this.lblremarks.Location = new System.Drawing.Point(0, 24);
            this.lblremarks.Name = "lblremarks";
            this.lblremarks.Size = new System.Drawing.Size(535, 24);
            this.lblremarks.Text = "备注";
            this.lblremarks.TextSize = new System.Drawing.Size(48, 14);
            // 
            // label1
            // 
            this.label1.AppearanceItemCaption.Options.UseTextOptions = true;
            this.label1.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.label1.Control = this.txtsupp;
            this.label1.Location = new System.Drawing.Point(535, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(301, 24);
            this.label1.Text = "供应商";
            this.label1.TextSize = new System.Drawing.Size(48, 14);
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.btnDel;
            this.layoutControlItem2.Location = new System.Drawing.Point(393, 48);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(269, 26);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.databind;
            this.layoutControlItem3.Location = new System.Drawing.Point(0, 74);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(836, 615);
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(0, 48);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(103, 26);
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(662, 48);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(174, 26);
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.btnAdd;
            this.layoutControlItem1.Location = new System.Drawing.Point(103, 48);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(290, 26);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // TestSampleList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1015, 762);
            this.Controls.Add(this.layoutControl1);
            this.MaximizeBox = false;
            this.Name = "TestSampleList";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "TestSampleList";
            ((System.ComponentModel.ISupportInitialize)(this.databind)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblsampletype)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblproductcode)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblqty)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lblremarks)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.label1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.TextBox txtqty;
        private System.Windows.Forms.TextBox txtsupp;
        private System.Windows.Forms.TextBox txtproductcode;
        private System.Windows.Forms.TextBox txtremarks;
        private System.Windows.Forms.TextBox txtsampletype;
        private DevExpress.XtraGrid.GridControl databind;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView;
        private DevExpress.XtraEditors.SimpleButton btnAdd;
        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraEditors.SimpleButton btnDel;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem lblsampletype;
        private DevExpress.XtraLayout.LayoutControlItem lblproductcode;
        private DevExpress.XtraLayout.LayoutControlItem lblqty;
        private DevExpress.XtraLayout.LayoutControlItem lblremarks;
        private DevExpress.XtraLayout.LayoutControlItem label1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
    }
}