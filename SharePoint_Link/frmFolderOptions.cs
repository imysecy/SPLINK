using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using SharePoint_Link.Utility;
using System.Xml;
using SharePoint_Link.UserModule;
using SharePoint_Link.Utility;
using Utility;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SharePoint_Link
{
    /// <summary>
    /// code of this class written by Joy
    /// </summary>
    public partial class frmFolderOptions : Form
    {
        public static Outlook._Application OutlookObj;
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder olInBox;
        Microsoft.Office.Interop.Outlook.MAPIFolder parentFolder;
        Outlook.MAPIFolder mappedFolder;
        public frmFolderOptions()
        {
            InitializeComponent();
        }
        /// <summary>
        /// code written by Joy
        /// on form load loads the folder names in the drop down list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmFolderOptions_Load(object sender, EventArgs e)
        {
            cmbOptions.Validating += new CancelEventHandler(cmbOptions_Validating);
           
            try
            {
                DataSet ds = new DataSet();
               
                if (File.Exists(UserLogManagerUtility.XMLFilePath))
                {
                    ds.Tables.Clear();
                    ds.ReadXml(UserLogManagerUtility.XMLFilePath);
                    try
                    {
                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            cmbOptions.Items.Insert(i, ds.Tables[0].Rows[i]["DisplayName"]);
                            //cmbOptions.Items.Add(ds.Tables[0].Rows[i]["DisplayName"]);
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            catch (Exception ex)
            {
            }

        }
        /// <summary>
        /// code written by Joy
        /// validates the folder names in the dropdownlist
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void cmbOptions_Validating(object sender, CancelEventArgs e)
        {
            ComboBox objComboBox = (ComboBox)sender;
            
            if (objComboBox.SelectedIndex < 0)
            {
                errorProvider1.SetError(objComboBox, "Select");
                e.Cancel = true;
            }
            else
            {
                errorProvider1.SetError(objComboBox, null);
            }
        }
        /// <summary>
        /// code written by Joy
        /// moves or copies mail items to the selected mapped folder
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOk_Click(object sender, EventArgs e)
        {
           
            if (this.ValidateChildren() == true)
            {
                if (ThisAddIn.IsUploadingFormIsOpen == true)
                {
                    if (Globals.ThisAddIn.frmlistObject != null)
                    {
                        Globals.ThisAddIn.frmlistObject.progressBar1.Value = Globals.ThisAddIn.frmlistObject.progressBar1.Minimum;
                        Globals.ThisAddIn.frmlistObject.lblPRStatus.Text = "";
                    }



                }
                string selected_mapiFolderName = cmbOptions.Text;
                this.Close();
                if (Globals.ThisAddIn.copy_button_clicked == true)
                {
                    OutlookObj = Globals.ThisAddIn.Application;
                    outlookNameSpace = OutlookObj.GetNamespace("MAPI");
                    olInBox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    parentFolder = (Microsoft.Office.Interop.Outlook.MAPIFolder)olInBox.Parent;
                    string mappedFolderName = cmbOptions.SelectedItem.ToString();
                    mappedFolder = MAPIFolderWrapper.GetFolder(parentFolder, mappedFolderName);
                    Outlook.MailItem mItem;
                    Outlook.MailItem copyMail;
                    Globals.ThisAddIn.isCopyRunninng = true;
                    foreach (Object obj in Globals.ThisAddIn.copySelected)
                    {
                        if (obj is Outlook.MailItem)
                        {
                            mItem = (Outlook.MailItem)obj;
                            copyMail = mItem.Copy() as Outlook.MailItem;
                            copyMail.Move(mappedFolder);
                            //doBackGroundUpload(mItem);
                        }
                    }
                }
                else if (Globals.ThisAddIn.move_button_clicked == true)
                {
                    
                        OutlookObj = Globals.ThisAddIn.Application;
                        outlookNameSpace = OutlookObj.GetNamespace("MAPI");
                        olInBox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                        parentFolder = (Microsoft.Office.Interop.Outlook.MAPIFolder)olInBox.Parent;
                        string mappedFolderName = cmbOptions.SelectedItem.ToString();
                        mappedFolder = MAPIFolderWrapper.GetFolder(parentFolder, mappedFolderName);
                        Outlook.MailItem mItem;
                        Globals.ThisAddIn.isMoveRunning = true;
                        foreach (Object obj in Globals.ThisAddIn.moveSelected)
                        {
                            if (obj is Outlook.MailItem)
                            {
                                mItem = (Outlook.MailItem)obj;
                                mItem.Move(mappedFolder);
                                //doBackGroundUpload(mItem);
                            }
                        }

                       
                    
                }
            }
            else
                return;
           
        }
        //private const int CP_NOCLOSE_BUTTON = 0x200;
        //protected override CreateParams CreateParams
        //{
        //    get
        //    {
        //        CreateParams myCp = base.CreateParams;
        //        myCp.ClassStyle = myCp.ClassStyle | CP_NOCLOSE_BUTTON;
        //        return myCp;
        //    }
        //}

      
    }

}
