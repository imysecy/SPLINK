﻿#pragma checksum "..\..\WPFBrowserUserControl.xaml" "{406ea660-64cf-4c82-b6f0-42d48172a799}" "E26D5B3EC20BB91E5F257E173CEB57E5"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.1
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms.Integration;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace SharePoint_Link {
    
    
    /// <summary>
    /// WPFBrowserUserControl
    /// </summary>
    public partial class WPFBrowserUserControl : System.Windows.Controls.UserControl, System.Windows.Markup.IComponentConnector {
        
        
        #line 15 "..\..\WPFBrowserUserControl.xaml"
        internal System.Windows.Controls.Button backButton;
        
        #line default
        #line hidden
        
        
        #line 16 "..\..\WPFBrowserUserControl.xaml"
        internal System.Windows.Controls.Button forwardButton;
        
        #line default
        #line hidden
        
        
        #line 17 "..\..\WPFBrowserUserControl.xaml"
        internal System.Windows.Controls.Button refreshButton;
        
        #line default
        #line hidden
        
        
        #line 19 "..\..\WPFBrowserUserControl.xaml"
        internal System.Windows.Controls.TextBox addressTextBox;
        
        #line default
        #line hidden
        
        
        #line 20 "..\..\WPFBrowserUserControl.xaml"
        internal System.Windows.Controls.Button goNavigateButton;
        
        #line default
        #line hidden
        
        
        #line 27 "..\..\WPFBrowserUserControl.xaml"
        internal System.Windows.Controls.Primitives.StatusBarItem informationStatusBarItem;
        
        #line default
        #line hidden
        
        
        #line 33 "..\..\WPFBrowserUserControl.xaml"
        internal System.Windows.Controls.WebBrowser webBrowser;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/SharePoint_Link;component/wpfbrowserusercontrol.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\WPFBrowserUserControl.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            
            #line 8 "..\..\WPFBrowserUserControl.xaml"
            ((SharePoint_Link.WPFBrowserUserControl)(target)).Loaded += new System.Windows.RoutedEventHandler(this.UserControl_Loaded);
            
            #line default
            #line hidden
            return;
            case 2:
            this.backButton = ((System.Windows.Controls.Button)(target));
            
            #line 15 "..\..\WPFBrowserUserControl.xaml"
            this.backButton.Click += new System.Windows.RoutedEventHandler(this.backButton_Click);
            
            #line default
            #line hidden
            return;
            case 3:
            this.forwardButton = ((System.Windows.Controls.Button)(target));
            
            #line 16 "..\..\WPFBrowserUserControl.xaml"
            this.forwardButton.Click += new System.Windows.RoutedEventHandler(this.forwardButton_Click);
            
            #line default
            #line hidden
            return;
            case 4:
            this.refreshButton = ((System.Windows.Controls.Button)(target));
            
            #line 17 "..\..\WPFBrowserUserControl.xaml"
            this.refreshButton.Click += new System.Windows.RoutedEventHandler(this.refreshButton_Click);
            
            #line default
            #line hidden
            return;
            case 5:
            this.addressTextBox = ((System.Windows.Controls.TextBox)(target));
            return;
            case 6:
            this.goNavigateButton = ((System.Windows.Controls.Button)(target));
            
            #line 20 "..\..\WPFBrowserUserControl.xaml"
            this.goNavigateButton.Click += new System.Windows.RoutedEventHandler(this.goNavigateButton_Click);
            
            #line default
            #line hidden
            return;
            case 7:
            this.informationStatusBarItem = ((System.Windows.Controls.Primitives.StatusBarItem)(target));
            return;
            case 8:
            this.webBrowser = ((System.Windows.Controls.WebBrowser)(target));
            
            #line 33 "..\..\WPFBrowserUserControl.xaml"
            this.webBrowser.Navigated += new System.Windows.Navigation.NavigatedEventHandler(this.webBrowser_Navigated);
            
            #line default
            #line hidden
            
            #line 33 "..\..\WPFBrowserUserControl.xaml"
            this.webBrowser.LoadCompleted += new System.Windows.Navigation.LoadCompletedEventHandler(this.webBrowser_LoadCompleted);
            
            #line default
            #line hidden
            
            #line 33 "..\..\WPFBrowserUserControl.xaml"
            this.webBrowser.Navigating += new System.Windows.Navigation.NavigatingCancelEventHandler(this.webBrowser_Navigating);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}
