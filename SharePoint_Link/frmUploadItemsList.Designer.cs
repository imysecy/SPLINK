namespace SharePoint_Link
{
    partial class frmUploadItemsList
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
     
        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dgvUploadImages = new System.Windows.Forms.DataGridView();
            this.colStatusImage = new System.Windows.Forms.DataGridViewImageColumn();
            this.colEdit = new System.Windows.Forms.DataGridViewLinkColumn();
            this.colCurrentStatus = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMailSubject = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colElapsedTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colFolderName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripSplitButtonClose = new System.Windows.Forms.ToolStripSplitButton();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.lblPleaseWaitMessage = new System.Windows.Forms.Label();
            this.lblTimeElapsed = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblPRStatus = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvUploadImages)).BeginInit();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvUploadImages
            // 
            this.dgvUploadImages.AllowUserToAddRows = false;
            this.dgvUploadImages.AllowUserToDeleteRows = false;
            this.dgvUploadImages.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvUploadImages.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvUploadImages.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dgvUploadImages.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dgvUploadImages.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvUploadImages.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colStatusImage,
            this.colEdit,
            this.colCurrentStatus,
            this.colMailSubject,
            this.colElapsedTime,
            this.colFolderName});
            this.dgvUploadImages.Location = new System.Drawing.Point(0, 1);
            this.dgvUploadImages.MinimumSize = new System.Drawing.Size(586, 200);
            this.dgvUploadImages.Name = "dgvUploadImages";
            this.dgvUploadImages.ReadOnly = true;
            this.dgvUploadImages.RowHeadersVisible = false;
            this.dgvUploadImages.RowTemplate.Height = 30;
            this.dgvUploadImages.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvUploadImages.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvUploadImages.Size = new System.Drawing.Size(784, 549);
            this.dgvUploadImages.TabIndex = 4;
            this.dgvUploadImages.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvUploadImages_CellClick);
            // 
            // colStatusImage
            // 
            this.colStatusImage.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.colStatusImage.FillWeight = 50F;
            this.colStatusImage.HeaderText = "";
            this.colStatusImage.Name = "colStatusImage";
            this.colStatusImage.ReadOnly = true;
            this.colStatusImage.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.colStatusImage.Width = 50;
            // 
            // colEdit
            // 
            this.colEdit.FillWeight = 69.22543F;
            this.colEdit.HeaderText = "";
            this.colEdit.Name = "colEdit";
            this.colEdit.ReadOnly = true;
            // 
            // colCurrentStatus
            // 
            this.colCurrentStatus.FillWeight = 138.4509F;
            this.colCurrentStatus.HeaderText = "Current Status";
            this.colCurrentStatus.Name = "colCurrentStatus";
            this.colCurrentStatus.ReadOnly = true;
            // 
            // colMailSubject
            // 
            this.colMailSubject.FillWeight = 69.22543F;
            this.colMailSubject.HeaderText = "Upload Item Name";
            this.colMailSubject.Name = "colMailSubject";
            this.colMailSubject.ReadOnly = true;
            // 
            // colElapsedTime
            // 
            this.colElapsedTime.FillWeight = 69.22543F;
            this.colElapsedTime.HeaderText = "ElapsedTime";
            this.colElapsedTime.Name = "colElapsedTime";
            this.colElapsedTime.ReadOnly = true;
            this.colElapsedTime.Visible = false;
            // 
            // colFolderName
            // 
            this.colFolderName.HeaderText = "FolderName";
            this.colFolderName.Name = "colFolderName";
            this.colFolderName.ReadOnly = true;
            this.colFolderName.Visible = false;
            // 
            // statusStrip1
            // 
            this.statusStrip1.BackColor = System.Drawing.Color.White;
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripSplitButtonClose,
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 536);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.statusStrip1.Size = new System.Drawing.Size(784, 26);
            this.statusStrip1.TabIndex = 6;
            this.statusStrip1.Text = "ITOPIA";
            // 
            // toolStripSplitButtonClose
            // 
            this.toolStripSplitButtonClose.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripSplitButtonClose.Image = global::SharePoint_Link.Properties.Resources.close;
            this.toolStripSplitButtonClose.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripSplitButtonClose.Name = "toolStripSplitButtonClose";
            this.toolStripSplitButtonClose.Size = new System.Drawing.Size(32, 24);
            this.toolStripSplitButtonClose.Text = "toolStripSplitButton1";
            this.toolStripSplitButtonClose.ButtonClick += new System.EventHandler(this.toolStripSplitButtonClose_ButtonClick);
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Font = new System.Drawing.Font("Times New Roman", 14F);
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(72, 21);
            this.toolStripStatusLabel1.Text = "ITOPIA";
            // 
            // lblPleaseWaitMessage
            // 
            this.lblPleaseWaitMessage.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.lblPleaseWaitMessage.AutoSize = true;
            this.lblPleaseWaitMessage.BackColor = System.Drawing.SystemColors.Window;
            this.lblPleaseWaitMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPleaseWaitMessage.Location = new System.Drawing.Point(323, 329);
            this.lblPleaseWaitMessage.Name = "lblPleaseWaitMessage";
            this.lblPleaseWaitMessage.Size = new System.Drawing.Size(178, 13);
            this.lblPleaseWaitMessage.TabIndex = 1;
            this.lblPleaseWaitMessage.Text = "Please Wait - Uploading Items";
            // 
            // lblTimeElapsed
            // 
            this.lblTimeElapsed.AutoSize = true;
            this.lblTimeElapsed.Location = new System.Drawing.Point(539, 544);
            this.lblTimeElapsed.Name = "lblTimeElapsed";
            this.lblTimeElapsed.Size = new System.Drawing.Size(74, 13);
            this.lblTimeElapsed.TabIndex = 7;
            this.lblTimeElapsed.Text = "Elapsed Time:";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(116, 396);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(422, 23);
            this.progressBar1.TabIndex = 8;
            // 
            // lblPRStatus
            // 
            this.lblPRStatus.AutoSize = true;
            this.lblPRStatus.Location = new System.Drawing.Point(133, 432);
            this.lblPRStatus.Name = "lblPRStatus";
            this.lblPRStatus.Size = new System.Drawing.Size(0, 13);
            this.lblPRStatus.TabIndex = 9;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.pictureBox1.Image = global::SharePoint_Link.Properties.Resources.wait;
            this.pictureBox1.Location = new System.Drawing.Point(359, 220);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(105, 144);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // frmUploadItemsList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.lblPRStatus);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.lblTimeElapsed);
            this.Controls.Add(this.lblPleaseWaitMessage);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.dgvUploadImages);
            this.Name = "frmUploadItemsList";
            this.Size = new System.Drawing.Size(784, 562);
            this.Resize += new System.EventHandler(this.frmUploadItemsList_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dgvUploadImages)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

       
       

        #endregion

        public System.Windows.Forms.DataGridView dgvUploadImages;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ToolStripSplitButton toolStripSplitButtonClose;
        private System.Windows.Forms.Label lblTimeElapsed;
        private System.Windows.Forms.DataGridViewImageColumn colStatusImage;
        private System.Windows.Forms.DataGridViewLinkColumn colEdit;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCurrentStatus;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMailSubject;
        private System.Windows.Forms.DataGridViewTextBoxColumn colElapsedTime;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFolderName;

        #endregion
        public System.Windows.Forms.ProgressBar progressBar1;
        public System.Windows.Forms.Label lblPRStatus;
        public System.Windows.Forms.Label lblPleaseWaitMessage;




    }
}
