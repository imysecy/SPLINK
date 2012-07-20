using System;
using System.Windows.Forms;

namespace SharePoint_Link
{
    /// <summary>
    /// <c>frmMessageWindow</c> inherits <c>Form</c>
    /// </summary>
    public partial class frmMessageWindow : Form
    {
        /// <summary>
        /// <c>frmMessageWindow</c> default constructor
        /// initializes component
        /// </summary>
        public frmMessageWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// <c>displayMessage</c> member field of type string
        /// holds the message to display message 
        /// </summary>
        private string displayMessage;


        /// <summary>
        /// <c>DisplayMessage</c> member property  
        /// encapsulates displayMessage
        /// </summary>
        public string DisplayMessage
        {
            get { return displayMessage; }
            set { displayMessage = value; }
        }


        /// <summary>
        /// <c>frmMessageWindow_Load</c> Event Handler
        /// display the message
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmMessageWindow_Load(object sender, EventArgs e)
        {
            try
            {

                if (!string.IsNullOrEmpty(displayMessage))
                {
                    lblMessage.Text = displayMessage;
                }
                this.TopLevel = true;
                this.TopMost = true;
                this.Activate();
            }
            catch { }
        }

        /// <summary>
        /// <c>btnOk_Click</c> Event Handler
        /// close the  message window form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }


    }
}