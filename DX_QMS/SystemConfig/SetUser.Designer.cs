namespace DX_QMS.SystemConfig
{
    partial class SetUser
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
            this.ribbon = new DevExpress.XtraBars.Ribbon.RibbonControl();
            this.btnSave = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtUserId = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cbGroup = new System.Windows.Forms.ComboBox();
            this.txtUserName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.cbDept2 = new System.Windows.Forms.ComboBox();
            this.cbDept = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtTel = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.cbBsGroup = new System.Windows.Forms.ComboBox();
            this.gbUser = new DevExpress.XtraEditors.GroupControl();
            ((System.ComponentModel.ISupportInitialize)(this.ribbon)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gbUser)).BeginInit();
            this.gbUser.SuspendLayout();
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
            this.ribbon.Size = new System.Drawing.Size(440, 54);
            this.ribbon.Toolbar.ShowCustomizeItem = false;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(115, 367);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 15;
            this.btnSave.Text = "保 存";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.ForeColor = System.Drawing.Color.Red;
            this.label6.Location = new System.Drawing.Point(155, 330);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(49, 14);
            this.label6.TabIndex = 16;
            this.label6.Text = "123456";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(41, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 14);
            this.label1.TabIndex = 0;
            this.label1.Text = "员工工号：";
            // 
            // txtUserId
            // 
            this.txtUserId.Location = new System.Drawing.Point(121, 29);
            this.txtUserId.Name = "txtUserId";
            this.txtUserId.Size = new System.Drawing.Size(120, 22);
            this.txtUserId.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(41, 70);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 14);
            this.label2.TabIndex = 1;
            this.label2.Text = "员工姓名：";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(34, 239);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(74, 14);
            this.label4.TabIndex = 3;
            this.label4.Text = "C/S权限组：";
            // 
            // cbGroup
            // 
            this.cbGroup.FormattingEnabled = true;
            this.cbGroup.Location = new System.Drawing.Point(121, 235);
            this.cbGroup.Name = "cbGroup";
            this.cbGroup.Size = new System.Drawing.Size(120, 22);
            this.cbGroup.TabIndex = 11;
            // 
            // txtUserName
            // 
            this.txtUserName.Location = new System.Drawing.Point(121, 65);
            this.txtUserName.Name = "txtUserName";
            this.txtUserName.Size = new System.Drawing.Size(120, 22);
            this.txtUserName.TabIndex = 6;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(41, 331);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(67, 14);
            this.label5.TabIndex = 12;
            this.label5.Text = "初始密码：";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(41, 199);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(67, 14);
            this.label8.TabIndex = 18;
            this.label8.Text = "派遣部门：";
            // 
            // cbDept2
            // 
            this.cbDept2.FormattingEnabled = true;
            this.cbDept2.Location = new System.Drawing.Point(121, 194);
            this.cbDept2.Name = "cbDept2";
            this.cbDept2.Size = new System.Drawing.Size(120, 22);
            this.cbDept2.TabIndex = 19;
            // 
            // cbDept
            // 
            this.cbDept.FormattingEnabled = true;
            this.cbDept.Location = new System.Drawing.Point(121, 151);
            this.cbDept.Name = "cbDept";
            this.cbDept.Size = new System.Drawing.Size(120, 22);
            this.cbDept.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(41, 156);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(67, 14);
            this.label3.TabIndex = 2;
            this.label3.Text = "所属部门：";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(41, 111);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(67, 14);
            this.label7.TabIndex = 17;
            this.label7.Text = "联系电话：";
            // 
            // txtTel
            // 
            this.txtTel.Location = new System.Drawing.Point(121, 106);
            this.txtTel.Name = "txtTel";
            this.txtTel.Size = new System.Drawing.Size(120, 22);
            this.txtTel.TabIndex = 20;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(34, 288);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(74, 14);
            this.label9.TabIndex = 21;
            this.label9.Text = "B/S权限组：";
            // 
            // cbBsGroup
            // 
            this.cbBsGroup.FormattingEnabled = true;
            this.cbBsGroup.Location = new System.Drawing.Point(121, 283);
            this.cbBsGroup.Name = "cbBsGroup";
            this.cbBsGroup.Size = new System.Drawing.Size(120, 22);
            this.cbBsGroup.TabIndex = 22;
            // 
            // gbUser
            // 
            this.gbUser.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.gbUser.Controls.Add(this.label5);
            this.gbUser.Controls.Add(this.btnSave);
            this.gbUser.Controls.Add(this.label6);
            this.gbUser.Controls.Add(this.label1);
            this.gbUser.Controls.Add(this.cbGroup);
            this.gbUser.Controls.Add(this.cbBsGroup);
            this.gbUser.Controls.Add(this.label9);
            this.gbUser.Controls.Add(this.label4);
            this.gbUser.Controls.Add(this.label2);
            this.gbUser.Controls.Add(this.txtUserId);
            this.gbUser.Controls.Add(this.txtUserName);
            this.gbUser.Controls.Add(this.cbDept2);
            this.gbUser.Controls.Add(this.label8);
            this.gbUser.Controls.Add(this.label7);
            this.gbUser.Controls.Add(this.txtTel);
            this.gbUser.Controls.Add(this.cbDept);
            this.gbUser.Controls.Add(this.label3);
            this.gbUser.Location = new System.Drawing.Point(73, 75);
            this.gbUser.Margin = new System.Windows.Forms.Padding(4);
            this.gbUser.Name = "gbUser";
            this.gbUser.Padding = new System.Windows.Forms.Padding(4);
            this.gbUser.Size = new System.Drawing.Size(287, 428);
            this.gbUser.TabIndex = 2;
            this.gbUser.Text = "用户维护";
            // 
            // SetUser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(440, 511);
            this.Controls.Add(this.gbUser);
            this.Controls.Add(this.ribbon);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "SetUser";
            this.Ribbon = this.ribbon;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "用户维护";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SetUser_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.ribbon)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gbUser)).EndInit();
            this.gbUser.ResumeLayout(false);
            this.gbUser.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraBars.Ribbon.RibbonControl ribbon;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtUserId;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cbGroup;
        private System.Windows.Forms.TextBox txtUserName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cbDept2;
        private System.Windows.Forms.ComboBox cbDept;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtTel;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox cbBsGroup;
        private DevExpress.XtraEditors.GroupControl gbUser;
    }
}