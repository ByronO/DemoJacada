using SHDocVw;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Navigation;
using WebBrowser = System.Windows.Controls.WebBrowser;

namespace DemoJacada
{

    public partial class MainWindow : Window
    {
        private WebBrowser webBrowser;

        public MainWindow()
        {
            InitializeComponent();
           
           
        }
        private void abrir_Click(object sender, RoutedEventArgs e)
        {
            WebBrowser wb = new WebBrowser();
            wb.Margin = new Thickness(122, 38, 30, 10);
            wb.Source = new Uri(txtUrl.Text);    //https://vivr.io/qz37X91
            wb.Visibility = Visibility.Visible;
         
            this.webBrowser = wb;
            this.webBrowser.LoadCompleted += new LoadCompletedEventHandler(browser_LoadCompleted);
            
            JacadaGrid.Children.Add(wb);
         
        }
        
        private void cerrar_Click(object sender, RoutedEventArgs e)
        {
            this.webBrowser.Dispose();
        }

        
        public void browser_LoadCompleted(object sender, NavigationEventArgs args)
        {
            try
            {
                FieldInfo webBrowserInfo = this.webBrowser.GetType().GetField("_axIWebBrowser2", BindingFlags.Instance | BindingFlags.NonPublic);

                object comWebBrowser = null;
                object zoomPercent = 60;
                if (webBrowserInfo != null)
                {
                    comWebBrowser = webBrowserInfo.GetValue(this.webBrowser);
                }
                if (comWebBrowser != null)
                {
                    InternetExplorer ie = (InternetExplorer)comWebBrowser;
                    ie.ExecWB(SHDocVw.OLECMDID.OLECMDID_OPTICAL_ZOOM, SHDocVw.OLECMDEXECOPT.OLECMDEXECOPT_PROMPTUSER, ref zoomPercent, IntPtr.Zero);
                }


            }
            catch (Exception ex)
            {

            }
        }
    }
}
