using System;
using System.Windows.Forms;

namespace SharePoint_Link
{
    /// <summary>
    /// <c>frmRename</c> class inherits <c>Form</c>
    /// this class implements  the functionality to rename  outlook  folder
    /// </summary>
    public partial class frmRename : Form
    {
        /// <summary>
        /// <c>frmRename</c> default constructor
        /// </summary>
        public frmRename()
        {
            InitializeComponent();
        }

        /// <summary>
        /// <c>btnRename_Click</c> Event handler
        /// close the window form
        /// not currently used
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRename_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// <c>btnCancel_Click</c> Event handler
        /// not currently used
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            txtNewFolderName.Text = string.Empty;
            this.Close();
        }
    }
}