namespace SharePoint_Link
{
    partial class frmConnectionProperties
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
            this.dgvUsersInformation = new System.Windows.Forms.DataGridView();
            this.Reconnect = new System.Windows.Forms.DataGridViewImageColumn();
            this.DisplayName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.URL = new System.Windows.Forms.DataGridViewLinkColumn();
            this.AuthenticationType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Status = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DateAdded = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LastUpdated = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OutlookLocation = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.GroupBox1 = new System.Windows.Forms.GroupBox();
            this.dtpLastUpdated = new System.Windows.Forms.DateTimePicker();
            this.txtDateUpdated = new System.Windows.Forms.TextBox();
            this.dtpAdded = new System.Windows.Forms.DateTimePicker();
            this.Label2 = new System.Windows.Forms.Label();
            this.txtLocation = new System.Windows.Forms.TextBox();
            this.txtDateAdded = new System.Windows.Forms.TextBox();
            this.btnAll = new System.Windows.Forms.Button();
            this.Label4 = new System.Windows.Forms.Label();
            this.Label3 = new System.Windows.Forms.Label();
            this.lblURL = new System.Windows.Forms.Label();
            this.lblDisplayName = new System.Windows.Forms.Label();
            this.txtDisplayName = new System.Windows.Forms.TextBox();
            this.txtURL = new System.Windows.Forms.TextBox();
            this.btnReset = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            ((System.ComponentModel.ISupportInitialize)(this.dgvUsersInformation)).BeginInit();
            this.GroupBox1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgvUsersInformation
            // 
            this.dgvUsersInformation.AllowUserToAddRows = false;
            this.dgvUsersInformation.AllowUserToDeleteRows = false;
            this.dgvUsersInformation.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvUsersInformation.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvUsersInformation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvUsersInformation.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Reconnect,
            this.DisplayName,
            this.URL,
            this.AuthenticationType,
            this.Status,
            this.DateAdded,
            this.LastUpdated,
            this.OutlookLocation});
            this.dgvUsersInformation.Location = new System.Drawing.Point(0, 153);
            this.dgvUsersInformation.Name = "dgvUsersInformation";
            this.dgvUsersInformation.ReadOnly = true;
            this.dgvUsersInformation.Size = new System.Drawing.Size(892, 395);
            this.dgvUsersInformation.TabIndex = 1;
            this.dgvUsersInformation.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvUsersInformation_CellClick);
            this.dgvUsersInformation.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvUsersInformation_CellContentClick);
            // 
            // Reconnect
            // 
            this.Reconnect.FillWeight = 59.24647F;
            this.Reconnect.HeaderText = "Reconnection";
            this.Reconnect.Image = global::SharePoint_Link.Properties.Resources.new_conn;
            this.Reconnect.Name = "Reconnect";
            this.Reconnect.ReadOnly = true;
            // 
            // DisplayName
            // 
            this.DisplayName.DataPropertyName = "DisplayName";
            this.DisplayName.FillWeight = 87.15306F;
            this.DisplayName.HeaderText = "Display Name Contains";
            this.DisplayName.Name = "DisplayName";
            this.DisplayName.ReadOnly = true;
            // 
            // URL
            // 
            this.URL.DataPropertyName = "URL";
            this.URL.FillWeight = 104.3531F;
            this.URL.HeaderText = "URL Contains";
            this.URL.Name = "URL";
            this.URL.ReadOnly = true;
            // 
            // AuthenticationType
            // 
            this.AuthenticationType.DataPropertyName = "AuthenticationType";
            this.AuthenticationType.FillWeight = 76.25893F;
            this.AuthenticationType.HeaderText = "Authentication ";
            this.AuthenticationType.Name = "AuthenticationType";
            this.AuthenticationType.ReadOnly = true;
            // 
            // Status
            // 
            this.Status.DataPropertyName = "Status";
            this.Status.FillWeight = 70F;
            this.Status.HeaderText = "Status";
            this.Status.Name = "Status";
            this.Status.ReadOnly = true;
            this.Status.Visible = false;
            // 
            // DateAdded
            // 
            this.DateAdded.DataPropertyName = "DateAdded";
            this.DateAdded.FillWeight = 76.25893F;
            this.DateAdded.HeaderText = "Date Added";
            this.DateAdded.Name = "DateAdded";
            this.DateAdded.ReadOnly = true;
            // 
            // LastUpdated
            // 
            this.LastUpdated.DataPropertyName = "LastUpload";
            this.LastUpdated.FillWeight = 76.25893F;
            this.LastUpdated.HeaderText = "Last Updated";
            this.LastUpdated.Name = "LastUpdated";
            this.LastUpdated.ReadOnly = true;
            // 
            // OutlookLocation
            // 
            this.OutlookLocation.DataPropertyName = "OutlookLocation";
            this.OutlookLocation.FillWeight = 76.25893F;
            this.OutlookLocation.HeaderText = "Outlook Location";
            this.OutlookLocation.Name = "OutlookLocation";
            this.OutlookLocation.ReadOnly = true;
            // 
            // GroupBox1
            // 
            this.GroupBox1.Controls.Add(this.dtpLastUpdated);
            this.GroupBox1.Controls.Add(this.txtDateUpdated);
            this.GroupBox1.Controls.Add(this.dtpAdded);
            this.GroupBox1.Controls.Add(this.Label2);
            this.GroupBox1.Controls.Add(this.txtLocation);
            this.GroupBox1.Controls.Add(this.txtDateAdded);
            this.GroupBox1.Controls.Add(this.btnAll);
            this.GroupBox1.Controls.Add(this.Label4);
            this.GroupBox1.Controls.Add(this.Label3);
            this.GroupBox1.Controls.Add(this.lblURL);
            this.GroupBox1.Controls.Add(this.lblDisplayName);
            this.GroupBox1.Controls.Add(this.txtDisplayName);
            this.GroupBox1.Controls.Add(this.txtURL);
            this.GroupBox1.Controls.Add(this.btnReset);
            this.GroupBox1.Controls.Add(this.btnUpdate);
            this.GroupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GroupBox1.Location = new System.Drawing.Point(12, 10);
            this.GroupBox1.Name = "GroupBox1";
            this.GroupBox1.Size = new System.Drawing.Size(868, 127);
            this.GroupBox1.TabIndex = 2;
            this.GroupBox1.TabStop = false;
            this.GroupBox1.Text = " Filter By";
            // 
            // dtpLastUpdated
            // 
            this.dtpLastUpdated.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpLastUpdated.Location = new System.Drawing.Point(569, 56);
            this.dtpLastUpdated.Name = "dtpLastUpdated";
            this.dtpLastUpdated.Size = new System.Drawing.Size(19, 20);
            this.dtpLastUpdated.TabIndex = 15;
            this.dtpLastUpdated.ValueChanged += new System.EventHandler(this.dtpLastUpdated_ValueChanged);
            // 
            // txtDateUpdated
            // 
            this.txtDateUpdated.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.txtDateUpdated.Location = new System.Drawing.Point(495, 56);
            this.txtDateUpdated.Name = "txtDateUpdated";
            this.txtDateUpdated.ReadOnly = true;
            this.txtDateUpdated.Size = new System.Drawing.Size(93, 20);
            this.txtDateUpdated.TabIndex = 17;
            // 
            // dtpAdded
            // 
            this.dtpAdded.CustomFormat = "";
            this.dtpAdded.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpAdded.Location = new System.Drawing.Point(569, 31);
            this.dtpAdded.Name = "dtpAdded";
            this.dtpAdded.Size = new System.Drawing.Size(18, 20);
            this.dtpAdded.TabIndex = 14;
            this.dtpAdded.ValueChanged += new System.EventHandler(this.dtpAdded_ValueChanged);
            // 
            // Label2
            // 
            this.Label2.AutoSize = true;
            this.Label2.Location = new System.Drawing.Point(604, 30);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(117, 13);
            this.Label2.TabIndex = 9;
            this.Label2.Text = "Location Contains :";
            // 
            // txtLocation
            // 
            this.txtLocation.Location = new System.Drawing.Point(725, 26);
            this.txtLocation.Name = "txtLocation";
            this.txtLocation.Size = new System.Drawing.Size(131, 20);
            this.txtLocation.TabIndex = 8;
            // 
            // txtDateAdded
            // 
            this.txtDateAdded.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.txtDateAdded.Location = new System.Drawing.Point(495, 30);
            this.txtDateAdded.Name = "txtDateAdded";
            this.txtDateAdded.ReadOnly = true;
            this.txtDateAdded.Size = new System.Drawing.Size(93, 20);
            this.txtDateAdded.TabIndex = 18;
            // 
            // btnAll
            // 
            this.btnAll.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAll.Location = new System.Drawing.Point(555, 94);
            this.btnAll.Name = "btnAll";
            this.btnAll.Size = new System.Drawing.Size(177, 23);
            this.btnAll.TabIndex = 16;
            this.btnAll.Text = "Show All Connections";
            this.btnAll.UseVisualStyleBackColor = true;
            this.btnAll.Click += new System.EventHandler(this.btnAll_Click);
            // 
            // Label4
            // 
            this.Label4.AutoSize = true;
            this.Label4.Location = new System.Drawing.Point(330, 59);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(163, 13);
            this.Label4.TabIndex = 11;
            this.Label4.Text = "Last Updated Date Range :";
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(370, 33);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(123, 13);
            this.Label3.TabIndex = 10;
            this.Label3.Text = "Date Added Range :";
            // 
            // lblURL
            // 
            this.lblURL.AutoSize = true;
            this.lblURL.Location = new System.Drawing.Point(66, 59);
            this.lblURL.Name = "lblURL";
            this.lblURL.Size = new System.Drawing.Size(40, 13);
            this.lblURL.TabIndex = 6;
            this.lblURL.Text = "URL :";
            // 
            // lblDisplayName
            // 
            this.lblDisplayName.AutoSize = true;
            this.lblDisplayName.Location = new System.Drawing.Point(14, 32);
            this.lblDisplayName.Name = "lblDisplayName";
            this.lblDisplayName.Size = new System.Drawing.Size(92, 13);
            this.lblDisplayName.TabIndex = 5;
            this.lblDisplayName.Text = "Display Name :";
            // 
            // txtDisplayName
            // 
            this.txtDisplayName.Location = new System.Drawing.Point(110, 30);
            this.txtDisplayName.Name = "txtDisplayName";
            this.txtDisplayName.Size = new System.Drawing.Size(208, 20);
            this.txtDisplayName.TabIndex = 3;
            // 
            // txtURL
            // 
            this.txtURL.Location = new System.Drawing.Point(110, 56);
            this.txtURL.Name = "txtURL";
            this.txtURL.Size = new System.Drawing.Size(208, 20);
            this.txtURL.TabIndex = 2;
            // 
            // btnReset
            // 
            this.btnReset.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnReset.Location = new System.Drawing.Point(362, 94);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(177, 23);
            this.btnReset.TabIndex = 1;
            this.btnReset.Text = "Reset Fields";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnUpdate.Location = new System.Drawing.Point(158, 94);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(177, 23);
            this.btnUpdate.TabIndex = 0;
            this.btnUpdate.Text = "Filter";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.BackColor = System.Drawing.Color.White;
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 536);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.statusStrip1.Size = new System.Drawing.Size(884, 26);
            this.statusStrip1.TabIndex = 7;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Font = new System.Drawing.Font("Times New Roman", 14F);
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(72, 21);
            this.toolStripStatusLabel1.Text = "ITOPIA";
            // 
            // frmConnectionProperties
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(884, 562);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.GroupBox1);
            this.Controls.Add(this.dgvUsersInformation);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(900, 600);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(900, 600);
            this.Name = "frmConnectionProperties";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Connection Properties";
            this.Load += new System.EventHandler(this.frmConnectionProperties_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvUsersInformation)).EndInit();
            this.GroupBox1.ResumeLayout(false);
            this.GroupBox1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.DataGridView dgvUsersInformation;
        internal System.Windows.Forms.GroupBox GroupBox1;
        internal System.Windows.Forms.DateTimePicker dtpLastUpdated;
        internal System.Windows.Forms.TextBox txtDateUpdated;
        internal System.Windows.Forms.DateTimePicker dtpAdded;
        internal System.Windows.Forms.TextBox txtDateAdded;
        internal System.Windows.Forms.Button btnAll;
        internal System.Windows.Forms.Label Label4;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.TextBox txtLocation;
        internal System.Windows.Forms.Label lblURL;
        internal System.Windows.Forms.Label lblDisplayName;
        internal System.Windows.Forms.TextBox txtDisplayName;
        internal System.Windows.Forms.TextBox txtURL;
        internal System.Windows.Forms.Button btnReset;
        internal System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.DataGridViewImageColumn Reconnect;
        private System.Windows.Forms.DataGridViewTextBoxColumn DisplayName;
        private System.Windows.Forms.DataGridViewLinkColumn URL;
        private System.Windows.Forms.DataGridViewTextBoxColumn AuthenticationType;
        private System.Windows.Forms.DataGridViewTextBoxColumn Status;
        private System.Windows.Forms.DataGridViewTextBoxColumn DateAdded;
        private System.Windows.Forms.DataGridViewTextBoxColumn LastUpdated;
        private System.Windows.Forms.DataGridViewTextBoxColumn OutlookLocation;


    }
}