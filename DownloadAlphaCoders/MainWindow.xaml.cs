using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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

namespace DownloadAlphaCoders
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public Notification notification = new Notification();
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = notification;

        }
        string title, folder = "a";
        ChromeDriver chromeDriver;
        List<ImageMetaData> lstRs = new List<ImageMetaData>();
        string[] stringSeparators = new string[] { "\r\n" };
        string[] stringSeparators2 = new string[] { "thumb-" };
        private void click1_Click(object sender, RoutedEventArgs e)
        {
            bool b = true;
            var thread = new Thread((ThreadStart)delegate
            {
                while (b)
                {

                    this.Dispatcher.Invoke(() =>
                    {
                        title = textTitle.Text;
                        folder = textFolder.Text;
                    });


                    notification.ActionNotifi = "Starting";
                    string currentLine = "";
                    currentLine = File.ReadAllText("abc.txt");

                    List<string> listStrLineElements = currentLine.Split(stringSeparators, StringSplitOptions.None).ToList();


                    ChromeDriverService service = ChromeDriverService.CreateDefaultService();
                    service.HideCommandPromptWindow = true;

                    notification.ActionNotifi = "Get list link picture";

                    var options = new ChromeOptions();
                    //options.AddArgument("headless");
                    chromeDriver = new ChromeDriver(service, options);

                    for (int i = 0; i < listStrLineElements.Count; i++)
                    {
                        try
                        {
                            ThreadCrawlData(listStrLineElements[i]);
                        }
                        catch (Exception)
                        {
                            continue;
                            throw;
                        }

                    } 
                       
                    




                    if (lstRs.Count > 0)
                    {
                        ExcelPackage excel = new ExcelPackage();

                        // name of the sheet 
                        var workSheet = excel.Workbook.Worksheets.Add("Sheet1");

                        // setting the properties 
                        // of the work sheet  
                        workSheet.TabColor = System.Drawing.Color.Black;
                        workSheet.DefaultRowHeight = 12;

                        // Setting the properties 
                        // of the first row 
                        workSheet.Row(1).Height = 20;
                        workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        workSheet.Row(1).Style.Font.Bold = true;

                        // Header of the Excel sheet 
                        // workSheet.Cells[1, 1].Value = "S.No";
                        workSheet.Cells[1, 1].Value = "Foldername";
                        workSheet.Cells[1, 2].Value = "Imagename";
                        workSheet.Cells[1, 3].Value = "Title";
                        workSheet.Cells[1, 4].Value = "Des";
                        workSheet.Cells[1, 5].Value = "Tag";
                        workSheet.Cells[1, 6].Value = "STT";

                        // Inserting the article data into excel 
                        // sheet by using the for each loop 
                        // As we have values to the first row  
                        // we will start with second row 
                        int recordIndex = 2;
                        notification.ActionNotifi = "Export excel";
                        foreach (var item in lstRs)
                        {
                            //workSheet.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                            workSheet.Cells[recordIndex, 1].Value = folder;
                            workSheet.Cells[recordIndex, 2].Value = item.Name;
                            workSheet.Cells[recordIndex, 3].Value = title;
                            workSheet.Cells[recordIndex, 4].Value = title;

                            workSheet.Cells[recordIndex, 5].Value = item.Tags;
                            workSheet.Cells[recordIndex, 6].Value = (recordIndex - 1).ToString();
                            workSheet.Cells[recordIndex, 7].Value = item.Url;
                            recordIndex++;
                        }

                        // By default, the column width is not  
                        // set to auto fit for the content 
                        // of the range, so we are using 
                        // AutoFit() method here.  
                        workSheet.Column(1).AutoFit();
                        workSheet.Column(2).AutoFit();
                        workSheet.Column(3).AutoFit();

                        // file name with .xlsx extension  
                        string path = Directory.GetParent(System.Reflection.Assembly.GetExecutingAssembly().Location).FullName;
                        string p_strPath = System.IO.Path.Combine(path, "listing.xlsx");

                        if (File.Exists(p_strPath))
                            File.Delete(p_strPath);

                        // Create excel file on physical disk  
                        FileStream objFileStrm = File.Create(p_strPath);
                        objFileStrm.Close();

                        // Write content to excel file  
                        File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
                        //Close Excel package 
                        excel.Dispose();

                        notification.ActionNotifi = "Done!!";
                        b = false;

                    }
                }
            });

            thread.Start();

        }

        private void ThreadCrawlData(string url)
        {

            chromeDriver.Url = url;
            chromeDriver.Navigate();
            //  Thread.Sleep(2000);
            IJavaScriptExecutor js = chromeDriver as IJavaScriptExecutor;
            var scriptTitle = "var a=$('.boxgrid >a:first-child').map(function() {return this.title;}).get(); return a;";
            var scriptSrc = "var a=$('.img-thumb').map(function() {return this.src;}).get();  return a;";
            var scriptHref = "var a=$('.boxgrid >a:first-child').map(function() {return this.href;}).get(); return a;";
            var lstTitle = (System.Collections.ObjectModel.ReadOnlyCollection<object>)js.ExecuteScript(scriptTitle);
            var lstUrl = (System.Collections.ObjectModel.ReadOnlyCollection<object>)js.ExecuteScript(scriptSrc);
            var lsttHref = (System.Collections.ObjectModel.ReadOnlyCollection<object>)js.ExecuteScript(scriptHref);
            notification.ActionNotifi = "Get metadata img:" + url;

            for (int i = 0; i < lstTitle.Count; i++)
            {
                ImageMetaData imageMeta = new ImageMetaData();
                

                string a = (string)lstUrl[i];
                string title = (string)lstTitle[i];

                chromeDriver.Url = (string)lsttHref[i];
                chromeDriver.Navigate();
                var scriptTitle2 = "var a=$('#list_tags>.well >a').map(function() {return this.text;}).get(); return a;";
                var lstTitle2 = (System.Collections.ObjectModel.ReadOnlyCollection<object>)js.ExecuteScript(scriptTitle2);

                List<string> lst = a.Split(stringSeparators2, StringSplitOptions.None).ToList();
                imageMeta.Name = lst[1];
                imageMeta.Url = a.Replace("thumb-", "");
                imageMeta.Tags =lstTitle2.Count>0 ? String.Join(", ", lstTitle2.ToArray()) : title.Replace(" ",",");

                lstRs.Add(imageMeta);
            }


        }

        


    }

    public class ImageMetaData
    {
        private string tags;
        private string url;
        private string name;

        public string Url { get => url; set => url = value; }
        public string Tags { get => tags; set => tags = value; }
        public string Name { get => name; set => name = value; }
    }

    public class Notification : INotifyPropertyChanged
    {
        protected string action;

        public string ActionNotifi
        {
            get { return action; }
            set
            {
                if (action != value)
                {
                    action = value;
                    OnPropertyChanged("ActionNotifi");

                }
            }
        }



        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyname)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyname));
            }

        }



    }
}
