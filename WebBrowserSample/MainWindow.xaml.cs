using System; // Uri
using System.IO;
using System.Windows; // Window, RoutedEventArgs
using System.Windows.Navigation; // NavigatingCancelEventArgs, NavigationEventArgs

namespace WebBrowserControlSample
{
    public partial class MainWindow : Window
    {
        ScriptableClass scriptableClass;

        public MainWindow()
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
                MessageBox.Show("Cannot go Forward. There needs to be something in the history to go forward to.");
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
                MessageBox.Show("Refresh what? You need to load a page in the web browser before you can refesh.");
            }
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
                MessageBox.Show("The Address URI must be absolute eg 'http://www.microsoft.com'");
                return;
            }
        }
        //private void goSourceButton_Click(object sender, RoutedEventArgs e)
        //{
        //    // Get URI to navigate to
        //    Uri uri = new Uri(this.addressTextBox.Text, UriKind.RelativeOrAbsolute);

        //    // Only absolute URIs can be navigated to.
        //    try
        //    {
        //        // Navigate to the desired URL by setting the .Source property
        //        this.webBrowser.Source = uri;
        //    }
        //    catch
        //    {
        //        MessageBox.Show("The Address URI must be absolute eg 'http://www.microsoft.com'");
        //        return;
        //    }
        //}
       

        
        //private void webBrowser_Navigating(object sender, NavigatingCancelEventArgs e)
        //{

        //    //string uriString = (e.Uri != null ? " to " + e.Uri : "");

        //    //// The WebBrowser control is about to locate and download the specified HTML document
        //    //this.informationStatusBarItem.Content = string.Format("Navigating{0}", uriString);

        //    //// Cancel navigation?
        //    //string msg = string.Format("Navigate{0}?", uriString);
        //    //MessageBoxResult result = MessageBox.Show(msg, "Navigate", MessageBoxButton.YesNo, MessageBoxImage.Question);
        //    //if (result == MessageBoxResult.No)
        //    //{
        //    //    e.Cancel = true;
        //    //    this.informationStatusBarItem.Content = string.Format("Canceled navigation to {0}", e.Uri);
        //    //}
        //}
        private void webBrowser_Navigated(object sender, NavigationEventArgs e)
        {
            // The WebBrowser control has located and begun downloading the specified HTML document
            string uriString = (e.Uri != null ? " to " + e.Uri : "");
            this.informationStatusBarItem.Content = string.Format("Navigated{0}", uriString);

            //if(e.Uri != null)
            //    addressTextBox.Text =  e.Uri.ToString();
        }
        private void webBrowser_LoadCompleted(object sender, NavigationEventArgs e)
        {
            // The WebBrowser control has completely downloaded the HTML document
            string uriString = (e.Uri != null ? " to " + e.Uri : "");
            this.informationStatusBarItem.Content = string.Format("Completed loading{0}", uriString);
            if(e.Uri != null)
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
                MessageBox.Show("Form Loaded", "Form Loaded", MessageBoxButton.OK, MessageBoxImage.None);
            //if (result == MessageBoxResult.No)
            //{
            //    e.Cancel = true;
            //    this.informationStatusBarItem.Content = string.Format("Canceled navigation to {0}", e.Uri);
            //}
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Get URI to navigate to
            Uri uri = new Uri(this.addressTextBox.Text, UriKind.RelativeOrAbsolute);

            // Only absolute URIs can be navigated to.
            try
            {
                // Navigate to the desired URL by setting the .Source property
                this.webBrowser.Source = uri;
            }
            catch
            {
                MessageBox.Show("The Address URI must be absolute eg 'http://www.microsoft.com'");
                return;
            }
        }

            
    }
}
