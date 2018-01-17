using CsQuery;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
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
using System.Collections.ObjectModel;

namespace AliParser
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            wb.Navigated += new NavigatedEventHandler(wb_Navigated);
            wb.LoadCompleted += new LoadCompletedEventHandler(wb_LoadCompleted);
        }
        ObservableCollection<Data> DataArray = new ObservableCollection<Data>();
        private void StartParseBtn_Click(object sender, RoutedEventArgs e)
        {
            if (DataArray.Count>0) {
                DataGrid.ItemsSource = DataArray;
            }
            String url = "";
            try
            {
                url = String.IsNullOrEmpty(UrlText.Text) ? throw (new ArgumentNullException()) : UrlText.Text;
            }
            catch
            {
                MessageBox.Show("Url is empty");
                return;
            }
            try
            {
                wb.Navigate(url);
            }
            catch
            {
                MessageBox.Show("Url is incorrect:" + UrlText.Text);
            }
        }

        void wb_LoadCompleted(object sender, NavigationEventArgs e)
        {
            try
            {
                dynamic doc = wb.Document;
                CQ cq = CQ.Create(doc.documentElement.InnerHtml);
            }
            catch
            {
                MessageBox.Show("Error in getting content in url"+ UrlText.Text, "Error", MessageBoxButton.OK,MessageBoxImage.Warning);
            }
        }
        void parsePage(CQ cq) {
            foreach (IDomObject dom in cq.Find("li.list-item")) {
            }
        }
#region Hack to avoid JS errors
        // to awoid JS errors
        void wb_Navigated(object sender, NavigationEventArgs e)
        {
            SetSilent(wb, true); // make it silent
        }
        public static void SetSilent(WebBrowser browser, bool silent)
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
#endregion
        #region Data table
        class Data
        {
            public Data(int Id, string Name, double PriceMin, double PriceMax, string Url)
            {
                this.Id = Id;
                this.Name = Name;
                this.PriceMin = PriceMin;
                this.PriceMax = PriceMax;
                this.Url = Url;
            }
            public int Id { get; set; }
            public string Name { get; set; }
            public double PriceMin { get; set; }
            public double PriceMax { get; set; }
            public string Url { get; set; }
        }
    }
#endregion
}
