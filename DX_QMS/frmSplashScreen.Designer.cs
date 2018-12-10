namespace DX_QMS
{
    partial class frmSplashScreen
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmSplashScreen));
            this.marqueeProgressBarControl1 = new DevExpress.XtraEditors.MarqueeProgressBarControl();
            this.labelControl = new DevExpress.XtraEditors.LabelControl();
            this.pictureEdit = new DevExpress.XtraEditors.PictureEdit();
            ((System.ComponentModel.ISupportInitialize)(this.marqueeProgressBarControl1.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureEdit.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // marqueeProgressBarControl1
            // 
            this.marqueeProgressBarControl1.EditValue = 0;
            this.marqueeProgressBarControl1.Location = new System.Drawing.Point(29, 413);
            this.marqueeProgressBarControl1.Name = "marqueeProgressBarControl1";
            this.marqueeProgressBarControl1.Properties.Appearance.BackColor = System.Drawing.Color.White;
            this.marqueeProgressBarControl1.Properties.LookAndFeel.SkinName = "VS2010";
            this.marqueeProgressBarControl1.Properties.LookAndFeel.UseDefaultLookAndFeel = false;
            this.marqueeProgressBarControl1.Size = new System.Drawing.Size(350, 11);
            this.marqueeProgressBarControl1.TabIndex = 5;
            // 
            // labelControl
            // 
            this.labelControl.Appearance.BackColor = System.Drawing.Color.White;
            this.labelControl.Appearance.Options.UseBackColor = true;
            this.labelControl.Location = new System.Drawing.Point(31, 393);
            this.labelControl.LookAndFeel.UseDefaultLookAndFeel = false;
            this.labelControl.Name = "labelControl";
            this.labelControl.Size = new System.Drawing.Size(55, 14);
            this.labelControl.TabIndex = 7;
            this.labelControl.Text = "Starting...";
            // 
            // pictureEdit
            // 
            this.pictureEdit.Cursor = System.Windows.Forms.Cursors.Default;
            this.pictureEdit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureEdit.EditValue = ((object)(resources.GetObject("pictureEdit.EditValue")));
            this.pictureEdit.Location = new System.Drawing.Point(0, 0);
            this.pictureEdit.Name = "pictureEdit";
            this.pictureEdit.Properties.ShowCameraMenuItem = DevExpress.XtraEditors.Controls.CameraMenuItemVisibility.Auto;
            this.pictureEdit.Properties.ZoomAccelerationFactor = 1D;
            this.pictureEdit.Size = new System.Drawing.Size(410, 500);
            this.pictureEdit.TabIndex = 8;
            // 
            // frmSplashScreen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(410, 500);
            this.Controls.Add(this.labelControl);
            this.Controls.Add(this.marqueeProgressBarControl1);
            this.Controls.Add(this.pictureEdit);
            this.Name = "frmSplashScreen";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.marqueeProgressBarControl1.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureEdit.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraEditors.MarqueeProgressBarControl marqueeProgressBarControl1;
        private DevExpress.XtraEditors.LabelControl labelControl;
        private DevExpress.XtraEditors.PictureEdit pictureEdit;
    }
}
