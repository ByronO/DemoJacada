using SHDocVw;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Navigation;


namespace DemoJacada
{

    public partial class MainWindow : Window
    {
        private System.Windows.Controls.WebBrowser webBrowser;

        public MainWindow()
        {
            InitializeComponent();
            this.webBrowser = new System.Windows.Controls.WebBrowser();
           
        }
        private void abrir_Click(object sender, RoutedEventArgs e)
        {
            this.webBrowser.Dispose();
            System.Windows.Controls.WebBrowser wb = new System.Windows.Controls.WebBrowser();
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
