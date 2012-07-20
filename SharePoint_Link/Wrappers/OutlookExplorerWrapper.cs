using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace SharePoint_Link
{
    /// <summary>
    /// Wrapper Class to handle the outlook Explorer
    /// We handle the outlook close event in this class
    /// </summary>
    public class OutlookExplorerWrapper
    {
        # region Global Varaibles

        private Outlook.Explorer m_WindowsNew;                         // wrapped window object
        private Office.CommandBars cbWidget;               // CommandBar

        # endregion

        #region Events

        public event EventHandler Close;
        public event EventHandler<InvalidateEventArgs> InvalidateControl;
        public event EventHandler MailitemCatchanged;

        #endregion

        #region Constructor

        /// <summary>
        /// <c>OutlookExplorerWrapper</c>
        /// Create new instance for explorer class 
        /// </summary>
        /// <param name="explorer"></param>
        public OutlookExplorerWrapper(Outlook.Explorer explorer)
        {
            //m_WindowsNew = explorer;
            m_WindowsNew = ThisAddIn.OutlookObj.ActiveExplorer();
            cbWidget = m_WindowsNew.CommandBars;
            ((Outlook.ExplorerEvents_Event)explorer).Close += new Microsoft.Office.Interop.Outlook.ExplorerEvents_CloseEventHandler(OutlookExplorerNew_Close);
        }




        /// <summary>
        /// <c>OutlookExplorerNew_Close</c> member function
        /// wrapper function to close the outlook window
        /// </summary>
        void OutlookExplorerNew_Close()
        {
            try
            {
                if (Close != null)
                {
                    Office.CommandBarPopup foundMenu = (Office.CommandBarPopup)cbWidget.FindControl(
                                Office.MsoControlType.msoControlPopup, System.Type.Missing, "ITOPIA Tools", true);
                    if (foundMenu != null)
                    {
                        foundMenu.Delete(System.Type.Missing);
                    }
                    Close(this, EventArgs.Empty);
                }
            }
            catch { }
            try
            {
                //Check OutLook open status. If it is open close the outlook.
                System.Diagnostics.Process[] outlookProcesses = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");
                if (outlookProcesses.Length > 0)
                {
                    outlookProcesses[0].Close();
                }
            }
            catch { }
        }

        #endregion

    }
}
