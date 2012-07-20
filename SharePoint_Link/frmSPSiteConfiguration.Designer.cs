namespace SharePoint_Link
{
    partial class frmSPSiteConfiguration
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
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUserName = new System.Windows.Forms.TextBox();
            this.rbtnManuallySpecified = new System.Windows.Forms.RadioButton();
            this.rbtnUseDomainCredentials = new System.Windows.Forms.RadioButton();
            this.txtURL = new System.Windows.Forms.TextBox();
            this.txtDisplayName = new System.Windows.Forms.TextBox();
            this.lblPassword = new System.Windows.Forms.Label();
            this.lblUserName = new System.Windows.Forms.Label();
            this.lblAuthentication = new System.Windows.Forms.Label();
            this.lblURL = new System.Windows.Forms.Label();
            this.lblDisplayName = new System.Windows.Forms.Label();
            this.lblversion = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnCancel.Location = new System.Drawing.Point(282, 211);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 75;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnOK.Location = new System.Drawing.Point(166, 211);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 74;
            this.btnOK.Text = "Ok";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(247, 170);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(134, 20);
            this.txtPassword.TabIndex = 72;
            this.txtPassword.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSPSiteConfiguration_KeyDown);
            // 
            // txtUserName
            // 
            this.txtUserName.Location = new System.Drawing.Point(247, 144);
            this.txtUserName.Name = "txtUserName";
            this.txtUserName.Size = new System.Drawing.Size(134, 20);
            this.txtUserName.TabIndex = 71;
            this.txtUserName.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSPSiteConfiguration_KeyDown);
            // 
            // rbtnManuallySpecified
            // 
            this.rbtnManuallySpecified.AutoSize = true;
            this.rbtnManuallySpecified.Checked = true;
            this.rbtnManuallySpecified.Location = new System.Drawing.Point(162, 119);
            this.rbtnManuallySpecified.Name = "rbtnManuallySpecified";
            this.rbtnManuallySpecified.Size = new System.Drawing.Size(114, 17);
            this.rbtnManuallySpecified.TabIndex = 70;
            this.rbtnManuallySpecified.TabStop = true;
            this.rbtnManuallySpecified.Text = "Manually Specified";
            this.rbtnManuallySpecified.UseVisualStyleBackColor = true;
            this.rbtnManuallySpecified.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSPSiteConfiguration_KeyDown);
            // 
            // rbtnUseDomainCredentials
            // 
            this.rbtnUseDomainCredentials.AutoSize = true;
            this.rbtnUseDomainCredentials.Location = new System.Drawing.Point(162, 99);
            this.rbtnUseDomainCredentials.Name = "rbtnUseDomainCredentials";
            this.rbtnUseDomainCredentials.Size = new System.Drawing.Size(138, 17);
            this.rbtnUseDomainCredentials.TabIndex = 69;
            this.rbtnUseDomainCredentials.Text = "Use Domain Credentials";
            this.rbtnUseDomainCredentials.UseVisualStyleBackColor = true;
            this.rbtnUseDomainCredentials.CheckedChanged += new System.EventHandler(this.rbtnUseDomainCredentials_CheckedChanged);
            this.rbtnUseDomainCredentials.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSPSiteConfiguration_KeyDown);
            // 
            // txtURL
            // 
            this.txtURL.BackColor = System.Drawing.Color.White;
            this.txtURL.Location = new System.Drawing.Point(162, 56);
            this.txtURL.Name = "txtURL";
            this.txtURL.Size = new System.Drawing.Size(272, 20);
            this.txtURL.TabIndex = 67;
            this.txtURL.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSPSiteConfiguration_KeyDown);
            // 
            // txtDisplayName
            // 
            this.txtDisplayName.Location = new System.Drawing.Point(162, 25);
            this.txtDisplayName.Name = "txtDisplayName";
            this.txtDisplayName.Size = new System.Drawing.Size(272, 20);
            this.txtDisplayName.TabIndex = 66;
            this.txtDisplayName.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSPSiteConfiguration_KeyDown);
            // 
            // lblPassword
            // 
            this.lblPassword.AutoSize = true;
            this.lblPassword.Location = new System.Drawing.Point(182, 173);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(59, 13);
            this.lblPassword.TabIndex = 80;
            this.lblPassword.Text = "Password :";
            // 
            // lblUserName
            // 
            this.lblUserName.AutoSize = true;
            this.lblUserName.Location = new System.Drawing.Point(178, 147);
            this.lblUserName.Name = "lblUserName";
            this.lblUserName.Size = new System.Drawing.Size(63, 13);
            this.lblUserName.TabIndex = 79;
            this.lblUserName.Text = "UserName :";
            // 
            // lblAuthentication
            // 
            this.lblAuthentication.AutoSize = true;
            this.lblAuthentication.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAuthentication.Location = new System.Drawing.Point(59, 99);
            this.lblAuthentication.Name = "lblAuthentication";
            this.lblAuthentication.Size = new System.Drawing.Size(97, 13);
            this.lblAuthentication.TabIndex = 78;
            this.lblAuthentication.Text = "Authentication :";
            // 
            // lblURL
            // 
            this.lblURL.AutoSize = true;
            this.lblURL.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblURL.Location = new System.Drawing.Point(13, 56);
            this.lblURL.Name = "lblURL";
            this.lblURL.Size = new System.Drawing.Size(143, 13);
            this.lblURL.TabIndex = 77;
            this.lblURL.Text = "Document Library URL :";
            // 
            // lblDisplayName
            // 
            this.lblDisplayName.AutoSize = true;
            this.lblDisplayName.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDisplayName.Location = new System.Drawing.Point(60, 28);
            this.lblDisplayName.Name = "lblDisplayName";
            this.lblDisplayName.Size = new System.Drawing.Size(96, 13);
            this.lblDisplayName.TabIndex = 76;
            this.lblDisplayName.Text = "Display  Name :";
            // 
            // lblversion
            // 
            this.lblversion.AutoSize = true;
            this.lblversion.Location = new System.Drawing.Point(163, 81);
            this.lblversion.Name = "lblversion";
            this.lblversion.Size = new System.Drawing.Size(44, 13);
            this.lblversion.TabIndex = 82;
            this.lblversion.Text = "version:";
            // 
            // frmSPSiteConfiguration
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(484, 262);
            this.ControlBox = false;
            this.Controls.Add(this.lblversion);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.txtPassword);
            this.Controls.Add(this.txtUserName);
            this.Controls.Add(this.rbtnManuallySpecified);
            this.Controls.Add(this.rbtnUseDomainCredentials);
            this.Controls.Add(this.txtURL);
            this.Controls.Add(this.txtDisplayName);
            this.Controls.Add(this.lblPassword);
            this.Controls.Add(this.lblUserName);
            this.Controls.Add(this.lblAuthentication);
            this.Controls.Add(this.lblURL);
            this.Controls.Add(this.lblDisplayName);
            this.MaximumSize = new System.Drawing.Size(500, 300);
            this.MinimumSize = new System.Drawing.Size(500, 300);
            this.Name = "frmSPSiteConfiguration";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Configure SharePoint Site";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.frmSPSiteConfiguration_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmSPSiteConfiguration_KeyDown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.TextBox txtUserName;
        private System.Windows.Forms.RadioButton rbtnManuallySpecified;
        private System.Windows.Forms.RadioButton rbtnUseDomainCredentials;
        private System.Windows.Forms.TextBox txtURL;
        private System.Windows.Forms.TextBox txtDisplayName;
        private System.Windows.Forms.Label lblPassword;
        private System.Windows.Forms.Label lblUserName;
        private System.Windows.Forms.Label lblAuthentication;
        private System.Windows.Forms.Label lblURL;
        private System.Windows.Forms.Label lblDisplayName;
        private System.Windows.Forms.Label lblversion;
    }
}