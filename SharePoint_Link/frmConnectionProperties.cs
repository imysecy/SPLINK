using System;
using System.Data;
using System.Text;
using System.Windows.Forms;
using SharePoint_Link.UserModule;
using SharePoint_Link.Utility;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Collections;

namespace SharePoint_Link
{
    /// <summary>
    /// <c>frmConnectionProperties</c> class inherits <c>Form</c>
    /// this class implements functions and properties  related to connection properties to 
    /// sharepoint mapped folders
    /// </summary>
    public partial class frmConnectionProperties : Form
    {
        # region Global Variables

        /// <summary>
        /// <c>dsUserInformation</c> member field of type  <c>DataSet</c>
        /// </summary>
        DataSet dsUserInformation = new DataSet();

        /// <summary>
        /// <c>addinExplorer</c> member field of type <c>Explorer</c>
        /// </summary>
        Outlook.Explorer addinExplorer;
        # endregion

        # region Constructor


        /// <summary>
        /// <c>frmConnectionProperties</c>  Default constructor
        /// </summary>
        public frmConnectionProperties()
        {
            InitializeComponent();
            dgvUsersInformation.AutoGenerateColumns = false;
        }

        # endregion

        # region From Events


        /// <summary>
        /// <c>frmConnectionProperties_Load</c> Event Handler
        /// it calls <c>BindGrid</c> member function to display connections id datagrid
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmConnectionProperties_Load(object sender, EventArgs e)
        {

            BindGrid();
        }


        /// <summary>
        /// <c>BindGrid</c> member function
        /// display all  sharepoint mapped connections 
        /// </summary>
        private void BindGrid()
        {
            try
            {
                if (File.Exists(UserLogManagerUtility.XMLFilePath))
                {

                    dsUserInformation.Tables.Clear();
                    dsUserInformation.ReadXml(UserLogManagerUtility.XMLFilePath);

                    dgvUsersInformation.DataSource = dsUserInformation.Tables[0];




                }
            }
            catch (Exception)
            {


            }
        }

        # endregion

        # region Button Events


        /// <summary>
        /// <c>btnUpdate_Click</c> Event Handler
        /// Filter the connection records based on display name, locaton etc
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (dsUserInformation.Tables.Count > 0)
            {
                dsUserInformation.Tables[0].DefaultView.RowFilter = "";
                DataView customeDataView = dsUserInformation.Tables[0].DefaultView;

                ArrayList criteriaList = new ArrayList();


                if (string.IsNullOrEmpty(txtDisplayName.Text) == false)
                {
                    criteriaList.Add("DisplayName LIKE '%" + txtDisplayName.Text + "%'");
                }
                if (string.IsNullOrEmpty(txtLocation.Text) == false)
                {
                    criteriaList.Add("OutlookLocation LIKE '%" + txtLocation.Text + "%'");
                }

                if (string.IsNullOrEmpty(txtURL.Text) == false)
                {
                    criteriaList.Add("URL LIKE '%" + txtURL.Text + "%'");
                }
                if (string.IsNullOrEmpty(txtDateAdded.Text) == false)
                {
                    criteriaList.Add("DateAdded LIKE '%" + txtDateAdded.Text + "%'");
                }
                if (string.IsNullOrEmpty(txtDateUpdated.Text) == false)
                {
                    criteriaList.Add("LastUpload LIKE '%" + txtDateUpdated.Text + "%'");
                }


                StringBuilder filterString = new StringBuilder();

                foreach (string criteria in criteriaList)
                {
                    if (string.IsNullOrEmpty(filterString.ToString()) == true)
                    {
                        filterString.Append(criteria);
                    }
                    else
                    {
                        filterString.Append(" AND ");
                        filterString.Append(criteria);
                    }
                }

                customeDataView.RowFilter = filterString.ToString();
                dgvUsersInformation.DataSource = customeDataView;
            }
        }


        /// <summary>
        /// <c>btnReset_Click</c> event Handler
        /// it calls <c>ResetAllFilterControls</c>  method to reset fields
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReset_Click(object sender, EventArgs e)
        {
            ResetAllFilterControls();
        }


        /// <summary>
        /// <c>btnAll_Click</c> Event Handler
        /// it calls <c>  ResetAllFilterControls </c> method to clear input fields and display all connections
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAll_Click(object sender, EventArgs e)
        {
            ResetAllFilterControls();
            if (dsUserInformation.Tables.Count > 0)
            {
                dsUserInformation.Tables[0].DefaultView.RowFilter = "";
                dgvUsersInformation.DataSource = dsUserInformation.Tables[0];
            }
        }

        # endregion

        # region DateTimePicker Events

        /// <summary>
        /// <c>dtpAdded_ValueChanged</c> event handler
        /// display the selected date in Date added Range field
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dtpAdded_ValueChanged(object sender, EventArgs e)
        {
            txtDateAdded.Text = dtpAdded.Value.ToShortDateString();
        }

        /// <summary>
        /// <c>dtpLastUpdated_ValueChanged</c> event handler
        /// it display selected date in last updated date range 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dtpLastUpdated_ValueChanged(object sender, EventArgs e)
        {
            txtDateUpdated.Text = dtpLastUpdated.Value.ToShortDateString();
        }



        # endregion

        # region Methods

        /// <summary>
        /// <c>ResetAllFilterControls</c> 
        /// Method to reset all the filter controls
        /// </summary>
        private void ResetAllFilterControls()
        {
            txtDisplayName.Text = String.Empty;
            txtLocation.Text = String.Empty;
            txtURL.Text = String.Empty;
            dtpAdded.Value = DateTime.Now;
            dtpLastUpdated.Value = DateTime.Now;
            txtDateAdded.Text = String.Empty;
            txtDateUpdated.Text = String.Empty;
        }

        # endregion

        /// <summary>
        /// <c>dgvUsersInformation_CellContentClick</c> Event Handler
        /// it display <c>frmSPSiteConfiguration</c> window form to update connection properties
        /// of the selected record
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvUsersInformation_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 0)
                {


                    string foldername = "";
                    foldername = dgvUsersInformation.Rows[e.RowIndex].Cells[1].Value.ToString();

                    frmSPSiteConfiguration frmSPSiteConfigurationObject = new frmSPSiteConfiguration();
                    frmSPSiteConfigurationObject.ShowEditForm(foldername);


                    if (frmSPSiteConfigurationObject.IsConfigureCompleted)
                    {
                        UserLogManagerUtility.UpdateFolderConfigDetails("", frmSPSiteConfigurationObject.FolderConfigProperties);
                        //foreach (MAPIFolderWrapper folder in myFolders)
                        //{
                        //    if (folder.FolderName == folderName)
                        //    {
                        //        folder.AttachedFolder.WebViewURL = frmSPSiteConfigurationObject.URL;
                        //        break;
                        //    }
                        //}

                        Microsoft.Office.Interop.Outlook.Application outlookObj = Globals.ThisAddIn.Application;
                        //Gte MAPI Name space
                        Microsoft.Office.Interop.Outlook.NameSpace outlookNameSpace = outlookObj.GetNamespace("MAPI");

                        Microsoft.Office.Interop.Outlook.MAPIFolder oInBox = outlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                        Microsoft.Office.Interop.Outlook.MAPIFolder parentFolder = (Microsoft.Office.Interop.Outlook.MAPIFolder)oInBox.Parent;
                        Microsoft.Office.Interop.Outlook.MAPIFolder f = MAPIFolderWrapper.GetFolder(parentFolder, foldername);
                        if (f != null)
                        {
                            f.WebViewURL = frmSPSiteConfigurationObject.URL;

                        }

                    }
                    BindGrid();



                }
            }
            catch (Exception ex)
            {


            }
        }

        /// <summary>
        /// <c>dgvUsersInformation_CellClick</c> event handler
        /// creates folder in outlook if it is present in config file but not in outlook
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvUsersInformation_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 0)
                {
                    string foldername = "";
                    foldername = dgvUsersInformation.Rows[e.RowIndex].Cells[1].Value.ToString();

                    frmSPSiteConfiguration frmSPSiteConfigurationObject = new frmSPSiteConfiguration();
                    frmSPSiteConfigurationObject.ShowEditForm(foldername);


                    if (frmSPSiteConfigurationObject.IsConfigureCompleted)
                    {
                        UserLogManagerUtility.UpdateFolderConfigDetails("", frmSPSiteConfigurationObject.FolderConfigProperties);

                        Microsoft.Office.Interop.Outlook.Application outlookObj = Globals.ThisAddIn.Application;
                        //Gte MAPI Name space
                        Microsoft.Office.Interop.Outlook.NameSpace outlookNameSpace = outlookObj.GetNamespace("MAPI");

                        Microsoft.Office.Interop.Outlook.MAPIFolder oInBox = outlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                        Microsoft.Office.Interop.Outlook.MAPIFolder parentFolder = (Microsoft.Office.Interop.Outlook.MAPIFolder)oInBox.Parent;
                        Microsoft.Office.Interop.Outlook.MAPIFolder f = MAPIFolderWrapper.GetFolder(parentFolder, foldername);

                        if (parentFolder.Name.Trim() != foldername.Trim())
                        {
                            if (f.Name.Trim() != parentFolder.Name.Trim())
                            {
                                if (f != null)
                                {

                                    f.WebViewURL = frmSPSiteConfigurationObject.URL;
                                    if (f.FolderPath.Contains("\\Deleted Items\\"))
                                    {
                                        try
                                        {
                                            f.Delete();
                                        }
                                        catch (Exception)
                                        {


                                        }
                                        CreateFolder(foldername, frmSPSiteConfigurationObject.FolderConfigProperties);
                                    }
                                }
                            }
                            else
                            {
                                CreateFolder(foldername, frmSPSiteConfigurationObject.FolderConfigProperties);
                            }
                        }


                    }


                    BindGrid();
                }
            }
            catch (Exception)
            {


            }
        }


        /// <summary>
        /// <c>CreateFolder</c> member function
        /// create folder  with provided  folder name and its properties
        /// </summary>
        /// <param name="foldername"></param>
        /// <param name="xMLLogProperties"></param>
        public void CreateFolder(string foldername, XMLLogProperties xMLLogProperties)
        {
            try
            {
                string path = UserLogManagerUtility.GetFolderOutLookLocation(foldername);

                path = path.Remove(0, 2);
                string[] folderpath = path.Split('\\');

                Microsoft.Office.Interop.Outlook.Application outlookObj = Globals.ThisAddIn.Application;
                addinExplorer = outlookObj.ActiveExplorer();

                Microsoft.Office.Interop.Outlook.NameSpace outlookNameSpace = outlookObj.GetNamespace("MAPI");
                Microsoft.Office.Interop.Outlook.MAPIFolder oInBox = outlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                Microsoft.Office.Interop.Outlook.MAPIFolder parentFolder = (Microsoft.Office.Interop.Outlook.MAPIFolder)oInBox.Parent;
                foreach (Outlook.Folder item in addinExplorer.Session.Folders)
                {
                    if (item.Name.Trim() == folderpath[0])
                    {
                        parentFolder = item;
                    }
                }
                //Gte MAPI Name space



                for (int i = 1; i < folderpath.Length; i++)
                {

                    Microsoft.Office.Interop.Outlook.MAPIFolder f = MAPIFolderWrapper.GetChildFolder(parentFolder, folderpath[i]);
                    if (f != null)
                    {

                        parentFolder = f;
                    }
                    else
                    {

                        if (i < folderpath.Length - 1)
                        {
                            Outlook.MAPIFolder newfolder;
                            newfolder = parentFolder.Folders.Add(folderpath[i], Type.Missing);
                            parentFolder = newfolder;
                        }
                        else
                        {
                            ThisAddIn tad = new ThisAddIn();

                            tad.ReConnection(xMLLogProperties, parentFolder);
                        }
                    }


                }

            }
            catch (Exception ex)
            {


            }
        }


    }
}