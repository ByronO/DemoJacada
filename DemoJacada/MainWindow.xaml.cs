using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DemoJacada
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
       

        public MainWindow()
        {
            InitializeComponent();
            loadURL();
        }
        

        private void loadURL()
        {
            JacadaBrowser.Source = new Uri("https://vivr.io/qz37X91");
            JacadaBrowser.Visibility = Visibility.Hidden;          
            
        }

        public void btnZoom_Click(object sender, RoutedEventArgs e) {            
            var wb = (dynamic)JacadaBrowser.GetType().GetField("_axIWebBrowser2",
              BindingFlags.Instance | BindingFlags.NonPublic)
              .GetValue(JacadaBrowser);

            wb.ExecWB(63, 2, 120, IntPtr.Zero);
            JacadaBrowser.Visibility = Visibility.Visible;
        }
    }
}
