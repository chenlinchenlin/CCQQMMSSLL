namespace DX_QMS.SystemConfig
{
    partial class UsersInfo
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
            this.btnSelect = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtId = new System.Windows.Forms.TextBox();
            this.txtName = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnSetPwd = new System.Windows.Forms.Button();
            this.groupBSid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.userGroupId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.deptid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupname = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.deptid2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.userDept = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tel = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.userName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.userId = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnReset = new System.Windows.Forms.Button();
            this.dgvUsers = new System.Windows.Forms.DataGridView();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.ribbon)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvUsers)).BeginInit();
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
            this.ribbon.Size = new System.Drawing.Size(884, 54);
            this.ribbon.Toolbar.ShowCustomizeItem = false;
            // 
            // btnSelect
            // 
            this.btnSelect.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnSelect.Location = new System.Drawing.Point(265, 212);
            this.btnSelect.Margin = new System.Windows.Forms.Padding(4);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(70, 27);
            this.btnSelect.TabIndex = 17;
            this.btnSelect.Text = "查 询";
            this.btnSelect.UseVisualStyleBackColor = true;
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(47, 43);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 14);
            this.label1.TabIndex = 0;
            this.label1.Text = "用户工号：";
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(459, 42);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 14);
            this.label2.TabIndex = 1;
            this.label2.Text = "用户姓名：";
            // 
            // txtId
            // 
            this.txtId.Location = new System.Drawing.Point(123, 40);
            this.txtId.Name = "txtId";
            this.txtId.Size = new System.Drawing.Size(125, 22);
            this.txtId.TabIndex = 2;
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(533, 38);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(125, 22);
            this.txtName.TabIndex = 3;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.txtId);
            this.groupBox1.Controls.Add(this.txtName);
            this.groupBox1.Location = new System.Drawing.Point(90, 106);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(703, 87);
            this.groupBox1.TabIndex = 23;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "查询条件";
            // 
            // btnSetPwd
            // 
            this.btnSetPwd.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnSetPwd.Location = new System.Drawing.Point(561, 212);
            this.btnSetPwd.Name = "btnSetPwd";
            this.btnSetPwd.Size = new System.Drawing.Size(70, 27);
            this.btnSetPwd.TabIndex = 21;
            this.btnSetPwd.Text = "清密码";
            this.btnSetPwd.UseVisualStyleBackColor = true;
            this.btnSetPwd.Click += new System.EventHandler(this.btnSetPwd_Click);
            // 
            // groupBSid
            // 
            this.groupBSid.DataPropertyName = "groupBSid";
            this.groupBSid.HeaderText = "B/S权限组";
            this.groupBSid.Name = "groupBSid";
            this.groupBSid.ReadOnly = true;
            this.groupBSid.Visible = false;
            // 
            // userGroupId
            // 
            this.userGroupId.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.userGroupId.DataPropertyName = "groupid";
            this.userGroupId.HeaderText = "权限组ID";
            this.userGroupId.Name = "userGroupId";
            this.userGroupId.ReadOnly = true;
            this.userGroupId.Visible = false;
            // 
            // deptid
            // 
            this.deptid.DataPropertyName = "deptid";
            this.deptid.HeaderText = "部门编码";
            this.deptid.Name = "deptid";
            this.deptid.ReadOnly = true;
            this.deptid.Visible = false;
            // 
            // groupname
            // 
            this.groupname.DataPropertyName = "groupname";
            this.groupname.HeaderText = "所属权限组";
            this.groupname.Name = "groupname";
            this.groupname.ReadOnly = true;
            // 
            // deptid2
            // 
            this.deptid2.DataPropertyName = "deptid2";
            this.deptid2.HeaderText = "派遣部门：";
            this.deptid2.Name = "deptid2";
            this.deptid2.ReadOnly = true;
            this.deptid2.Visible = false;
            // 
            // userDept
            // 
            this.userDept.DataPropertyName = "deptname";
            this.userDept.HeaderText = "所属部门";
            this.userDept.Name = "userDept";
            this.userDept.ReadOnly = true;
            // 
            // tel
            // 
            this.tel.DataPropertyName = "tel";
            this.tel.HeaderText = "联系电话";
            this.tel.Name = "tel";
            this.tel.ReadOnly = true;
            // 
            // userName
            // 
            this.userName.DataPropertyName = "username";
            this.userName.HeaderText = "用户姓名";
            this.userName.Name = "userName";
            this.userName.ReadOnly = true;
            // 
            // userId
            // 
            this.userId.DataPropertyName = "userid";
            this.userId.HeaderText = "用户工号";
            this.userId.Name = "userId";
            this.userId.ReadOnly = true;
            // 
            // btnReset
            // 
            this.btnReset.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnReset.Location = new System.Drawing.Point(191, 212);
            this.btnReset.Margin = new System.Windows.Forms.Padding(4);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(70, 27);
            this.btnReset.TabIndex = 16;
            this.btnReset.Text = "重 置";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // dgvUsers
            // 
            this.dgvUsers.AllowUserToAddRows = false;
            this.dgvUsers.AllowUserToDeleteRows = false;
            this.dgvUsers.AllowUserToOrderColumns = true;
            this.dgvUsers.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.dgvUsers.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvUsers.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            this.dgvUsers.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvUsers.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.userId,
            this.userName,
            this.tel,
            this.userDept,
            this.deptid2,
            this.groupname,
            this.deptid,
            this.userGroupId,
            this.groupBSid});
            this.dgvUsers.Location = new System.Drawing.Point(90, 267);
            this.dgvUsers.MultiSelect = false;
            this.dgvUsers.Name = "dgvUsers";
            this.dgvUsers.ReadOnly = true;
            this.dgvUsers.RowHeadersVisible = false;
            this.dgvUsers.RowTemplate.Height = 23;
            this.dgvUsers.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvUsers.Size = new System.Drawing.Size(703, 315);
            this.dgvUsers.TabIndex = 22;
            this.dgvUsers.DoubleClick += new System.EventHandler(this.btnUpdate_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnDelete.Location = new System.Drawing.Point(487, 212);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(70, 27);
            this.btnDelete.TabIndex = 20;
            this.btnDelete.Text = "删 除";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnUpdate.Location = new System.Drawing.Point(413, 212);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(70, 27);
            this.btnUpdate.TabIndex = 19;
            this.btnUpdate.Text = "修 改";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnAdd.Location = new System.Drawing.Point(339, 212);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(70, 27);
            this.btnAdd.TabIndex = 18;
            this.btnAdd.Text = "新 增";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // UsersInfo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 615);
            this.Controls.Add(this.btnSelect);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnSetPwd);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.dgvUsers);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.ribbon);
            this.Name = "UsersInfo";
            this.Ribbon = this.ribbon;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "用户信息";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.UsersInfo_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ribbon)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvUsers)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraBars.Ribbon.RibbonControl ribbon;
        private System.Windows.Forms.Button btnSelect;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtId;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnSetPwd;
        private System.Windows.Forms.DataGridViewTextBoxColumn groupBSid;
        private System.Windows.Forms.DataGridViewTextBoxColumn userGroupId;
        private System.Windows.Forms.DataGridViewTextBoxColumn deptid;
        private System.Windows.Forms.DataGridViewTextBoxColumn groupname;
        private System.Windows.Forms.DataGridViewTextBoxColumn deptid2;
        private System.Windows.Forms.DataGridViewTextBoxColumn userDept;
        private System.Windows.Forms.DataGridViewTextBoxColumn tel;
        private System.Windows.Forms.DataGridViewTextBoxColumn userName;
        private System.Windows.Forms.DataGridViewTextBoxColumn userId;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.DataGridView dgvUsers;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Button btnAdd;
    }
}