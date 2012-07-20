using System;

namespace SPOutLook
{
    /// <summary>
    /// <c>SPCopyClass</c>
    /// Not in use
    /// </summary>
    public class SPCopyClass
    {
        SPCopyService.Copy copyws;

        public SPCopyClass()
        {
            copyws = new SPCopyService.Copy();
            copyws.CopyIntoItemsCompleted += new SPCopyService.CopyIntoItemsCompletedEventHandler(copyws_CopyIntoItemsCompleted);
            copyws.GetItemCompleted += new SPCopyService.GetItemCompletedEventHandler(listService_GetListItemsCompleted);
        }


        public void copyws_CopyIntoItemsCompleted(object sender, SPCopyService.CopyIntoItemsCompletedEventArgs e)
        {
           
            throw new NotImplementedException();
        }


        public delegate void GetItemCompletedEventHandler(object sender, SPCopyService.GetItemCompletedEventArgs e);
        public event GetItemCompletedEventHandler GetItemCompleted;

        public void listService_GetListItemsCompleted(object sender, SPCopyService.GetItemCompletedEventArgs e)
        {
           
        }


        public void CopyIntoItemsAsync()
        {

        }
    }
}
