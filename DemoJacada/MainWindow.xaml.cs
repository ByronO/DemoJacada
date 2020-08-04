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


        public MainWindow()
        {
            InitializeComponent();
            JacadaBrowser.LoadCompleted += new LoadCompletedEventHandler(browser_LoadCompleted);
            JacadaBrowser.Navigated += new NavigatedEventHandler(wbMain_Navigated);
        }
        private void abrir_Click(object sender, RoutedEventArgs e)
        {

            JacadaBrowser.Navigate(new Uri("https://vivr.io/qz37X91"));
            //JacadaBrowser.Source = new Uri("https://vivr.io/qz37X91");
            JacadaBrowser.Visibility = Visibility.Visible;
        }

        private void cerrar_Click(object sender, RoutedEventArgs e)
        {
            JacadaBrowser.Visibility = Visibility.Hidden;
        }

        public void browser_LoadCompleted(object sender, NavigationEventArgs args)
        {
            try
            {
                FieldInfo webBrowserInfo = JacadaBrowser.GetType().GetField("_axIWebBrowser2", BindingFlags.Instance | BindingFlags.NonPublic);

                object comWebBrowser = null;
                object zoomPercent = 60;
                if (webBrowserInfo != null)
                {
                    comWebBrowser = webBrowserInfo.GetValue(JacadaBrowser);
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



        void wbMain_Navigated(object sender, NavigationEventArgs e)
        {
            SetSilent(JacadaBrowser, false); // make it silent
        }


        public static void SetSilent(WebBrowser browser, Boolean silent)
        {
            if (browser == null)
                throw new ArgumentNullException("browser");

            // get an IWebBrowser2 from the document
            IOleServiceProvider sp = browser.Document as IOleServiceProvider;
            if (sp != null)
            {
                Guid IID_IWebBrowserApp = new Guid("0002DF05-0000-0000-C000-000000000046");
                Guid IID_IWebBrowser2 = new Guid("D30C1661-CDAF-11d0-8A3E-00C04FC9E26E");

                object webBrowser;
                sp.QueryService(ref IID_IWebBrowserApp, ref IID_IWebBrowser2, out webBrowser);
                if (webBrowser != null)
                {
                    webBrowser.GetType().InvokeMember("Silent", BindingFlags.Instance | BindingFlags.Public | BindingFlags.PutDispProperty, null, webBrowser, new object[] { silent });
                }
            }
        }


        [ComImport, Guid("6D5140C1-7436-11CE-8034-00AA006009FA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleServiceProvider
        {
            [PreserveSig]
            int QueryService([In] ref Guid guidService, [In] ref Guid riid, [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);
        }
    }
}
