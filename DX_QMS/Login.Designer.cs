namespace DX_QMS
{
    partial class Login
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Login));
            this.cboxRemember = new System.Windows.Forms.CheckBox();
            this.pictureBox_min = new System.Windows.Forms.PictureBox();
            this.pictureBox_max = new System.Windows.Forms.PictureBox();
            this.login_loading = new DevExpress.XtraSplashScreen.SplashScreenManager(this, typeof(global::DevExpress.BarManager.LoadingForm), true, true);
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.txtuser = new DevExpress.XtraEditors.ComboBoxEdit();
            this.txtpassword = new DevExpress.XtraEditors.TextEdit();
            this.btnlogin = new DevExpress.XtraEditors.SimpleButton();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_min)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_max)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtuser.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtpassword.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // cboxRemember
            // 
            this.cboxRemember.AutoSize = true;
            this.cboxRemember.BackColor = System.Drawing.Color.Transparent;
            this.cboxRemember.ForeColor = System.Drawing.Color.Gray;
            this.cboxRemember.Location = new System.Drawing.Point(67, 206);
            this.cboxRemember.Name = "cboxRemember";
            this.cboxRemember.Size = new System.Drawing.Size(74, 18);
            this.cboxRemember.TabIndex = 27;
            this.cboxRemember.Text = "记住密码";
            this.cboxRemember.UseVisualStyleBackColor = false;
            // 
            // pictureBox_min
            // 
            this.pictureBox_min.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox_min.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox_min.BackgroundImage")));
            this.pictureBox_min.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox_min.Location = new System.Drawing.Point(346, 20);
            this.pictureBox_min.Margin = new System.Windows.Forms.Padding(1);
            this.pictureBox_min.Name = "pictureBox_min";
            this.pictureBox_min.Size = new System.Drawing.Size(16, 16);
            this.pictureBox_min.TabIndex = 28;
            this.pictureBox_min.TabStop = false;
            this.pictureBox_min.Tag = "0";
            this.pictureBox_min.Click += new System.EventHandler(this.pictureBox_min_Click);
            this.pictureBox_min.MouseEnter += new System.EventHandler(this.pictureBox_min_MouseEnter);
            this.pictureBox_min.MouseLeave += new System.EventHandler(this.pictureBox_min_MouseLeave);
            // 
            // pictureBox_max
            // 
            this.pictureBox_max.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox_max.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox_max.BackgroundImage")));
            this.pictureBox_max.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox_max.Location = new System.Drawing.Point(376, 20);
            this.pictureBox_max.Name = "pictureBox_max";
            this.pictureBox_max.Size = new System.Drawing.Size(16, 16);
            this.pictureBox_max.TabIndex = 29;
            this.pictureBox_max.TabStop = false;
            this.pictureBox_max.Tag = "1";
            this.pictureBox_max.Click += new System.EventHandler(this.pictureBox_min_Click);
            this.pictureBox_max.MouseEnter += new System.EventHandler(this.pictureBox_min_MouseEnter);
            this.pictureBox_max.MouseLeave += new System.EventHandler(this.pictureBox_min_MouseLeave);
            // 
            // login_loading
            // 
            this.login_loading.ClosingDelay = 500;
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.BackColor = System.Drawing.Color.Transparent;
            this.linkLabel1.LinkColor = System.Drawing.Color.FromArgb(((int)(((byte)(102)))), ((int)(((byte)(153)))), ((int)(((byte)(0)))));
            this.linkLabel1.Location = new System.Drawing.Point(306, 208);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(35, 14);
            this.linkLabel1.TabIndex = 30;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "重 置";
            this.linkLabel1.Click += new System.EventHandler(this.btnreset_Click);
            // 
            // txtuser
            // 
            this.txtuser.Location = new System.Drawing.Point(108, 114);
            this.txtuser.Name = "txtuser";
            this.txtuser.Properties.Appearance.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtuser.Properties.Appearance.Options.UseFont = true;
            this.txtuser.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.txtuser.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.txtuser.Size = new System.Drawing.Size(233, 24);
            this.txtuser.TabIndex = 31;
            this.txtuser.SelectedValueChanged += new System.EventHandler(this.txtuser_SelectedValueChanged);
            // 
            // txtpassword
            // 
            this.txtpassword.Location = new System.Drawing.Point(108, 161);
            this.txtpassword.Name = "txtpassword";
            this.txtpassword.Properties.Appearance.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtpassword.Properties.Appearance.Options.UseFont = true;
            this.txtpassword.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.txtpassword.Properties.PasswordChar = '#';
            this.txtpassword.Size = new System.Drawing.Size(233, 24);
            this.txtpassword.TabIndex = 32;
            this.txtpassword.KeyUp += new System.Windows.Forms.KeyEventHandler(this.txtpassword_KeyUp);
            // 
            // btnlogin
            // 
            this.btnlogin.Appearance.BackColor = System.Drawing.Color.OliveDrab;
            this.btnlogin.Appearance.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnlogin.Appearance.ForeColor = System.Drawing.Color.White;
            this.btnlogin.Appearance.Options.UseBackColor = true;
            this.btnlogin.Appearance.Options.UseFont = true;
            this.btnlogin.Appearance.Options.UseForeColor = true;
            this.btnlogin.AppearancePressed.BackColor = System.Drawing.Color.DarkOliveGreen;
            this.btnlogin.AppearancePressed.Options.UseBackColor = true;
            this.btnlogin.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.btnlogin.Location = new System.Drawing.Point(63, 238);
            this.btnlogin.LookAndFeel.SkinMaskColor = System.Drawing.Color.DarkSeaGreen;
            this.btnlogin.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat;
            this.btnlogin.LookAndFeel.UseDefaultLookAndFeel = false;
            this.btnlogin.Name = "btnlogin";
            this.btnlogin.Size = new System.Drawing.Size(280, 34);
            this.btnlogin.TabIndex = 33;
            this.btnlogin.Text = "登   录";
            this.btnlogin.Click += new System.EventHandler(this.btnlogin_Click);
            // 
            // Login
            // 
            this.Appearance.BackColor = System.Drawing.Color.DimGray;
            this.Appearance.Options.UseBackColor = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImageLayoutStore = System.Windows.Forms.ImageLayout.None;
            this.BackgroundImageStore = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImageStore")));
            this.ClientSize = new System.Drawing.Size(410, 332);
            this.Controls.Add(this.btnlogin);
            this.Controls.Add(this.txtpassword);
            this.Controls.Add(this.txtuser);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.pictureBox_max);
            this.Controls.Add(this.pictureBox_min);
            this.Controls.Add(this.cboxRemember);
            this.DoubleBuffered = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.LookAndFeel.UseDefaultLookAndFeel = false;
            this.Name = "Login";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Login_Form";
            this.TransparencyKey = System.Drawing.Color.DimGray;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Login_Form_FormClosed);
            this.Load += new System.EventHandler(this.Login_Form_Load);
            this.Shown += new System.EventHandler(this.Login_Form_Shown);
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Login_MouseDown);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_min)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_max)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtuser.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtpassword.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.CheckBox cboxRemember;
        private System.Windows.Forms.PictureBox pictureBox_min;
        private System.Windows.Forms.PictureBox pictureBox_max;
        private DevExpress.XtraSplashScreen.SplashScreenManager login_loading;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private DevExpress.XtraEditors.ComboBoxEdit txtuser;
        private DevExpress.XtraEditors.TextEdit txtpassword;
        private DevExpress.XtraEditors.SimpleButton btnlogin;
    }
}