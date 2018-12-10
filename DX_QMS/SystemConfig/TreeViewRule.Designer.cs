namespace DX_QMS.SystemConfig
{
    partial class TreeViewRule
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
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("权限管理");
            this.ribbon = new DevExpress.XtraBars.Ribbon.RibbonControl();
            this.lblIsWeb = new System.Windows.Forms.Label();
            this.cbIsWeb = new System.Windows.Forms.ComboBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cbGroupID = new System.Windows.Forms.ComboBox();
            this.treeRule = new System.Windows.Forms.TreeView();
            ((System.ComponentModel.ISupportInitialize)(this.ribbon)).BeginInit();
            this.SuspendLayout();
            // 
            // ribbon
            // 
            this.ribbon.ExpandCollapseItem.Id = 0;
            this.ribbon.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.ribbon.ExpandCollapseItem});
            this.ribbon.Location = new System.Drawing.Point(0, 0);
            this.ribbon.MaxItemId = 1;
            this.ribbon.Name = "ribbon";
            this.ribbon.ShowToolbarCustomizeItem = false;
            this.ribbon.Size = new System.Drawing.Size(678, 54);
            this.ribbon.Toolbar.ShowCustomizeItem = false;
            // 
            // lblIsWeb
            // 
            this.lblIsWeb.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.lblIsWeb.AutoSize = true;
            this.lblIsWeb.Location = new System.Drawing.Point(86, 82);
            this.lblIsWeb.Name = "lblIsWeb";
            this.lblIsWeb.Size = new System.Drawing.Size(74, 14);
            this.lblIsWeb.TabIndex = 12;
            this.lblIsWeb.Text = "B/S权限组：";
            // 
            // cbIsWeb
            // 
            this.cbIsWeb.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.cbIsWeb.FormattingEnabled = true;
            this.cbIsWeb.Items.AddRange(new object[] {
            "N",
            "Y"});
            this.cbIsWeb.Location = new System.Drawing.Point(176, 79);
            this.cbIsWeb.Name = "cbIsWeb";
            this.cbIsWeb.Size = new System.Drawing.Size(69, 22);
            this.cbIsWeb.TabIndex = 11;
            this.cbIsWeb.SelectedIndexChanged += new System.EventHandler(this.cbIsWeb_SelectedIndexChanged);
            // 
            // btnSave
            // 
            this.btnSave.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btnSave.Location = new System.Drawing.Point(517, 79);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 10;
            this.btnSave.Text = "保 存";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(251, 82);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 14);
            this.label1.TabIndex = 9;
            this.label1.Text = "选择权限组：";
            // 
            // cbGroupID
            // 
            this.cbGroupID.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.cbGroupID.FormattingEnabled = true;
            this.cbGroupID.Location = new System.Drawing.Point(348, 79);
            this.cbGroupID.Name = "cbGroupID";
            this.cbGroupID.Size = new System.Drawing.Size(163, 22);
            this.cbGroupID.TabIndex = 8;
            this.cbGroupID.SelectedIndexChanged += new System.EventHandler(this.cbGroupID_SelectedIndexChanged);
            // 
            // treeRule
            // 
            this.treeRule.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.treeRule.CheckBoxes = true;
            this.treeRule.Location = new System.Drawing.Point(89, 115);
            this.treeRule.Name = "treeRule";
            treeNode1.Name = "gen";
            treeNode1.Text = "权限管理";
            this.treeRule.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode1});
            this.treeRule.ShowLines = false;
            this.treeRule.ShowRootLines = false;
            this.treeRule.Size = new System.Drawing.Size(503, 473);
            this.treeRule.TabIndex = 7;
            this.treeRule.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.treeRule_AfterCheck);
            this.treeRule.AfterExpand += new System.Windows.Forms.TreeViewEventHandler(this.treeRule_AfterExpand);
            // 
            // TreeViewRule
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(678, 616);
            this.Controls.Add(this.lblIsWeb);
            this.Controls.Add(this.cbIsWeb);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbGroupID);
            this.Controls.Add(this.treeRule);
            this.Controls.Add(this.ribbon);
            this.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "TreeViewRule";
            this.Ribbon = this.ribbon;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "权限管理";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.TreeViewRule_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ribbon)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraBars.Ribbon.RibbonControl ribbon;
        private System.Windows.Forms.Label lblIsWeb;
        private System.Windows.Forms.ComboBox cbIsWeb;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbGroupID;
        private System.Windows.Forms.TreeView treeRule;
    }
}