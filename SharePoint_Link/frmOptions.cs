using System;
using System.Windows.Forms;
using SharePoint_Link.UserModule;

namespace SharePoint_Link
{

    /// <summary>
    /// <c>frmOptions</c> class inherits <c>Form</c>
    /// implements  the functionality to display and update Auto Delete Email option
    /// </summary>
    public partial class frmOptions : Form
    {
        /// <summary>
        /// <c>valueUpdated</c> member field of type string
        /// </summary>
        bool valueUpdated = false;


        /// <summary>
        /// <c>ValuesUpdated</c>  member property 
        /// encapsulates valueUpdated member field
        /// </summary>
        public bool ValuesUpdated
        {
            get { return valueUpdated; }
        }

        /// <summary>
        /// <c>isAutoDeleteChecked</c> member property of type bool
        /// holds  the information whether the autodelete option is checked or not
        /// </summary>
        public bool isAutoDeleteChecked
        {
            get { return autoDeleteEmails.Checked; }
        }

        /// <summary>
        /// <c>frmOptions</c> constructor
        /// gets autodelete  value from configuration file and display  in form
        /// </summary>
        /// <param name="userOptions"></param>
        public frmOptions(XMLLogOptions userOptions)
        {
            InitializeComponent();

            autoDeleteEmails.Checked = userOptions.AutoDeleteEmails;
        }

        /// <summary>
        /// <c>checkBox1_CheckedChanged</c> Event handler 
        /// not currently used
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// <c>btnOK_Click</c> Event Handler
        /// sets the member field valueUpdated = true and hide the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOK_Click(object sender, EventArgs e)
        {

            if (btnOK.Text.ToUpper() == "OK")
            {
                valueUpdated = true;
                this.Hide();
                //UpdateFolderConfigrationDetails(true);
            }
            else
            {
                //UpdateFolderConfigrationDetails(false);
            }
        }


        /// <summary>
        ///<c>btnCancel_Click</c>  cancel event handler
        /// set member field valueUpdated=false and hide the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            //IsConfigureCompleted = false;
            valueUpdated = false;
            this.Hide();
        }
    }
}
