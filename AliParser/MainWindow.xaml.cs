using AngleSharp.Dom.Html;
using AngleSharp.Parser.Html;
using AngleSharp;
using AngleSharp.Dom;
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
using System.Text.RegularExpressions;

namespace AliParser
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        ObservableCollection<Data> DataArray = new ObservableCollection<Data>();
        public MainWindow()
        {
            InitializeComponent();
            DataGrid.ItemsSource = DataArray;
        }
        private async void StartParseBtn_Click(object sender, RoutedEventArgs e)
        {
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
                IDocument htmlDocument = await getHtml(UrlText.Text);
                parsePage(htmlDocument);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Url is incorrect:" + UrlText.Text);
            }
        }

        //async void wb_LoadCompleted(object sender, NavigationEventArgs e)
        //{
        //    try
        //    {

        //    }
        //    catch
        //    {
        //        MessageBox.Show("Error in getting content in url" + UrlText.Text, "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
        //    }
        //}
        async System.Threading.Tasks.Task<IDocument> getHtml(string url)
        {
            var htmlParser = new HtmlParser();
            var config = Configuration.Default.WithDefaultLoader();
            try
            {
                IDocument htmlDocument = await BrowsingContext.New(config).OpenAsync(url);
                return htmlDocument;
            }
            catch
            {
                MessageBox.Show("Не удалось открыть страницу, проверьте правильность ссылки");
                return null;
            }
        }
        async void parsePage(IDocument htmlDocument)
        {
            foreach (IElement item in htmlDocument.QuerySelectorAll("#list-items li"))
            {
                string price = item.QuerySelector(".price span.value")?.TextContent;
                List<double> prices = new List<double>();
                foreach (Match m in Regex.Matches(price, "[0-9,]*"))
                {
                    double dPrice = 0;
                    Double.TryParse(m.Value, out dPrice);
                    if (dPrice > 0)
                    {
                        prices.Add(dPrice);
                    }
                }
                string name = item.QuerySelector(".product")?.GetAttribute("title");
                string rate = item.QuerySelector(".star")?.GetAttribute("title");
                double dRate = 0;
                if (!String.IsNullOrEmpty(rate))
                {
                    foreach (Match m in Regex.Matches(rate, "[0-9.]*"))
                    {
                        if (!String.IsNullOrEmpty(m.Value) && dRate == 0)
                        {
                            Double.TryParse(m.Value.Replace(".", ","), out dRate);
                            break;
                        }
                    }
                }
                string reviews = String.IsNullOrEmpty(item.QuerySelector(".rate-num")?.GetAttribute("title"))
                    ? item.QuerySelector(".rate-num")?.TextContent
                    : item.QuerySelector(".rate-num")?.GetAttribute("title");
                int dReviews = 0;
                if (!String.IsNullOrEmpty(reviews))
                {
                    foreach (Match m in Regex.Matches(reviews, "[0-9]*"))
                    {
                        if (!String.IsNullOrEmpty(m.Value))
                            Int32.TryParse(m.Value, out dReviews);
                    }
                }
                string orders = item.QuerySelector(".order-num-a em")?.TextContent;

                int dOrders = 0;

                if (!String.IsNullOrEmpty(orders))
                {
                    foreach (Match m in Regex.Matches(orders, "[0-9]*"))
                    {
                        if (!String.IsNullOrEmpty(m.Value))
                            Int32.TryParse(m.Value, out dOrders);
                    }
                }
                string url = "https:" + item.QuerySelector(".product").GetAttribute("href");
                if (prices.Count == 1)
                {
                    DataArray.Add(new Data(DataArray.Count, name, dRate, dOrders, dReviews, url, prices[0]));
                }
                else if (prices.Count == 2)
                {
                    DataArray.Add(new Data(DataArray.Count, name, dRate, dOrders, dReviews, url, prices[0], prices[1]));
                }
            }
            if (!String.IsNullOrEmpty(htmlDocument.QuerySelector(".page-next")?.GetAttribute("href")))
            {
                IDocument htmlDocumentNext = await getHtml("http:" + htmlDocument.QuerySelector(".page-next")?.GetAttribute("href"));
                parsePage(htmlDocumentNext);
            }
            else
            {
                return;
            }
        }
        private void ExportToExcel()
        {
            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "ExportedFromDatGrid";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;

                //Loop through each row and read value from each column. 
                for (int i = 0; i < DataGrid.Items.Count - 1; i++)
                {
                    for (int j = 0; j < DataGrid.Columns.Count; j++)
                    {
                        // Excel index starts from 1,1. As first Row would have the Column headers, adding a condition check. 
                        if (cellRowIndex == 1)
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = DataGrid.Columns[j].Header;
                        }
                        else
                        {
                            switch (j){
                                case 0:
                                    worksheet.Cells[cellRowIndex, cellColumnIndex] = DataArray[i].Id;
                                    break;
                                case 1:
                                    worksheet.Cells[cellRowIndex, cellColumnIndex] = DataArray[i].Name;
                                    break;
                                case 2:
                                    worksheet.Cells[cellRowIndex, cellColumnIndex] = DataArray[i].PriceMin;
                                    break;
                                case 3:
                                    worksheet.Cells[cellRowIndex, cellColumnIndex] = DataArray[i].PriceMax;
                                    break;
                                case 4:
                                    worksheet.Cells[cellRowIndex, cellColumnIndex] = DataArray[i].Rate;
                                    break;
                                case 5:
                                    worksheet.Cells[cellRowIndex, cellColumnIndex] = DataArray[i].Orders;
                                    break;
                                case 6:
                                    worksheet.Cells[cellRowIndex, cellColumnIndex] = DataArray[i].Reviews;
                                    break;
                                case 7:
                                    worksheet.Cells[cellRowIndex, cellColumnIndex] = DataArray[i].Url;
                                    break;
                            }
                        }
                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                Microsoft.Win32.SaveFileDialog saveDialog = new Microsoft.Win32.SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 2;

                Nullable<bool> result = saveDialog.ShowDialog();
                if (result == true)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Export Successful");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
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
            public Data(int Id, string Name, double Rate, int Orders, int Reviews, string Url, double PriceMin, double PriceMax = 0)
            {
                this.Id = Id;
                this.Name = Name;
                this.PriceMin = PriceMin;
                this.PriceMax = PriceMax;
                this.Rate = Rate;
                this.Orders = Orders;
                this.Reviews = Reviews;
                this.Url = Url;
            }
            public int Id { get; set; }
            public string Name { get; set; }
            public double PriceMin { get; set; }
            public double PriceMax { get; set; }
            public double Rate { get; set; }
            public int Orders { get; set; }
            public int Reviews { get; set; }
            public string Url { get; set; }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ExportToExcel();
        }
    }
    #endregion
}
