namespace DX_QMS
{
    partial class TestValueQtyList
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblunit = new System.Windows.Forms.Label();
            this.lblsamplefactqty = new System.Windows.Forms.Label();
            this.lblrsno = new System.Windows.Forms.Label();
            this.lblsampleqty = new System.Windows.Forms.Label();
            this.btnclear = new DevExpress.XtraEditors.SimpleButton();
            this.btnsave = new DevExpress.XtraEditors.SimpleButton();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.AutoScroll = true;
            this.panel1.Location = new System.Drawing.Point(81, 117);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1035, 403);
            this.panel1.TabIndex = 11;
            // 
            // lblunit
            // 
            this.lblunit.AutoSize = true;
            this.lblunit.Location = new System.Drawing.Point(1071, 30);
            this.lblunit.Name = "lblunit";
            this.lblunit.Size = new System.Drawing.Size(0, 14);
            this.lblunit.TabIndex = 6;
            // 
            // lblsamplefactqty
            // 
            this.lblsamplefactqty.AutoSize = true;
            this.lblsamplefactqty.Location = new System.Drawing.Point(1022, 30);
            this.lblsamplefactqty.Name = "lblsamplefactqty";
            this.lblsamplefactqty.Size = new System.Drawing.Size(0, 14);
            this.lblsamplefactqty.TabIndex = 7;
            // 
            // lblrsno
            // 
            this.lblrsno.AutoSize = true;
            this.lblrsno.Location = new System.Drawing.Point(975, 59);
            this.lblrsno.Name = "lblrsno";
            this.lblrsno.Size = new System.Drawing.Size(0, 14);
            this.lblrsno.TabIndex = 8;
            // 
            // lblsampleqty
            // 
            this.lblsampleqty.AutoSize = true;
            this.lblsampleqty.Location = new System.Drawing.Point(975, 30);
            this.lblsampleqty.Name = "lblsampleqty";
            this.lblsampleqty.Size = new System.Drawing.Size(0, 14);
            this.lblsampleqty.TabIndex = 9;
            // 
            // btnclear
            // 
            this.btnclear.Appearance.Font = new System.Drawing.Font("Tahoma", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnclear.Appearance.Options.UseFont = true;
            this.btnclear.Location = new System.Drawing.Point(461, 559);
            this.btnclear.Name = "btnclear";
            this.btnclear.Size = new System.Drawing.Size(75, 23);
            this.btnclear.TabIndex = 13;
            this.btnclear.Text = "清除";
            this.btnclear.Click += new System.EventHandler(this.btnclear_Click);
            // 
            // btnsave
            // 
            this.btnsave.Appearance.Font = new System.Drawing.Font("Tahoma", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnsave.Appearance.Options.UseFont = true;
            this.btnsave.Location = new System.Drawing.Point(692, 559);
            this.btnsave.Name = "btnsave";
            this.btnsave.Size = new System.Drawing.Size(69, 23);
            this.btnsave.TabIndex = 12;
            this.btnsave.Text = "确定";
            this.btnsave.Click += new System.EventHandler(this.btnsave_Click);
            // 
            // TestValueQtyList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(1205, 707);
            this.Controls.Add(this.btnsave);
            this.Controls.Add(this.btnclear);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.lblunit);
            this.Controls.Add(this.lblsamplefactqty);
            this.Controls.Add(this.lblrsno);
            this.Controls.Add(this.lblsampleqty);
            this.Name = "TestValueQtyList";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "测试值";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.TestValueQtyList_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblunit;
        private System.Windows.Forms.Label lblsamplefactqty;
        private System.Windows.Forms.Label lblrsno;
        private System.Windows.Forms.Label lblsampleqty;
        private DevExpress.XtraEditors.SimpleButton btnclear;
        private DevExpress.XtraEditors.SimpleButton btnsave;
    }
}