using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SharePoint_Link
{
    /// <summary>
    /// Interaction logic for WPFBrowserUserControl.xaml
    /// </summary>
    public partial class WPFBrowserUserControl : UserControl
    {
        public WPFBrowserUserControl()
        {
            InitializeComponent();
        }
        private void backButton_Click(object sender, RoutedEventArgs e)
        {
            // Navigate to the previous HTML document, if there is one
            if (this.webBrowser.CanGoBack)
            {
                this.webBrowser.GoBack();
            }
            else
            {
                MessageBox.Show("Cannot go back. There needs to be something in the history to go back to.");
            }
        }
        private void forwardButton_Click(object sender, RoutedEventArgs e)
        {
            // Navigate to the next HTML document, if there is one
            if (this.webBrowser.CanGoForward)
            {
                this.webBrowser.GoForward();
            }
            else
            {
                MessageBox.Show("Cannot go forward. There needs to be something in the history to go forward to.");
            }
        }
        private void refreshButton_Click(object sender, RoutedEventArgs e)
        {
            if (webBrowser.IsLoaded && webBrowser.Source != null)
            {
                this.webBrowser.Refresh();
            }
            else
            {
                MessageBox.Show("Cannot Refresh. You need to load a page in the web browser before you can refesh.");
            }
        }

        public void setNavigationUrl(string newUrl)
        {
            addressTextBox.Text = newUrl;
        }

        private void goNavigateButton_Click(object sender, RoutedEventArgs e)
        {
            // Get URI to navigate to
            Uri uri = new Uri(this.addressTextBox.Text, UriKind.RelativeOrAbsolute);

            // Only absolute URIs can be navigated to.
            try
            {
                // Navigate to the desired URL by calling the .Navigate method
                this.webBrowser.Navigate(uri);
            }
            catch
            {
                MessageBox.Show("The Address URI must be absolute eg 'http://www.myitopia.com'");
                return;
            }
        }
        private void webBrowser_Navigated(object sender, NavigationEventArgs e)
        {
            // The WebBrowser control has located and begun downloading the specified HTML document
            string uriString = (e.Uri != null ? " to " + e.Uri : "");
            this.informationStatusBarItem.Content = string.Format("Navigated{0}", uriString);
        }
        private void webBrowser_LoadCompleted(object sender, NavigationEventArgs e)
        {
            // The WebBrowser control has completely downloaded the HTML document
            string uriString = (e.Uri != null ? " to " + e.Uri : "");
            this.informationStatusBarItem.Content = string.Format("Completed loading{0}", uriString);
            if (e.Uri != null)
                addressTextBox.Text = e.Uri.ToString();

        }

        private void webBrowser_Navigating(object sender, NavigatingCancelEventArgs e)
        {
            string uriString = (e.Uri != null ? " to " + e.Uri : "");

            //The WebBrowser control is about to locate and download the specified HTML document
            this.informationStatusBarItem.Content = string.Format("Navigating{0}", uriString);
            // Cancel navigation?

            string msg = string.Format("Navigate{0}?", uriString);
            if (msg.Contains("java"))
                MessageBox.Show("Item Loaded", "Item Loaded", MessageBoxButton.OK, MessageBoxImage.None);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            // Get URI to navigate to
            this.addressTextBox.Text = ThisAddIn.CurrentWebUrlLink;
            Uri uri = new Uri(this.addressTextBox.Text, UriKind.RelativeOrAbsolute);

            // Only absolute URIs can be navigated to.
            try
            {
                // Navigate to the desired URL by setting the .Source property
                this.webBrowser.Source = uri;
            }
            catch
            {
                MessageBox.Show("The Address URI must be absolute eg 'http://www.myitopia.com'");
                return;
            }
        }


    }
}
