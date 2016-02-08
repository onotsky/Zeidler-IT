using System;
using System.Collections.Generic;
using System.Linq;
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
using Microsoft.Win32;
using System.Data;
using GemBox.Spreadsheet;
using HtmlAgilityPack;
using System.IO;

namespace Html2Excel
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private HtmlDocument hDoc;
        public MainWindow()
        {
            InitializeComponent();
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            hDoc = new HtmlDocument();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            string htmlType = string.Empty;

            OpenFileDialog oFile = new OpenFileDialog();
            oFile.Filter = "HTM-Dateien|*.htm|HTML-Dateien|*.html";
            Nullable<bool> result = oFile.ShowDialog();

            if (result == true)
            {
                using (StreamReader sRead = new StreamReader(oFile.FileName))
                {
                    List<string> strList = new List<string>();
                    string str = string.Empty;
                    int count = 0;

                    while ((str = sRead.ReadLine()) != null)
                    {
                        if (str.Contains("\0!\0D\0O\0C\0T\0Y\0P\0E"))
                        {
                            htmlType = str.Replace("\0", string.Empty);
                        }
                        else if (str.Contains("\0<\0h\0e\0a\0d\0>\0") && count != 0)
                        {
                            using (StreamWriter sWrite = new StreamWriter(File.Open(string.Format("{0}DAT{1}.htm", AppDomain.CurrentDomain.BaseDirectory, count.ToString()), FileMode.Create), Encoding.UTF8))
                            {
                                sWrite.WriteLine(htmlType);
                                sWrite.WriteLine("<html>");
                                foreach (string _str in strList)
                                {
                                    sWrite.WriteLine(_str);
                                }
                                sWrite.WriteLine("</html>");
                            }
                            strList.Clear();
                            strList.Add(str.Replace("\0", string.Empty));
                            count++;
                        }
                        else if (str.Contains("\0<\0h\0e\0a\0d\0>\0") && count == 0)
                        {
                            count++;
                            strList.Add(str.Replace("\0", string.Empty));
                        }
                        else if (String.IsNullOrWhiteSpace(str))
                        {
                            continue;
                        }
                        else if (str.Equals("\0"))
                        {
                            continue;
                        }
                        else if (str.Contains("\0<\0h\0t\0m\0l\0>") || str.Contains("\0<\0/\0h\0t\0m\0l\0"))
                        {
                            continue;
                        }
                        else if(str.Contains("\0<\0h\0r\0>"))
                        {
                            continue;
                        }
                        else
                        {
                            strList.Add(str.Replace("\0", string.Empty));
                        }
                    }
                }

            }



        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog oFile = new OpenFileDialog();
            oFile.Filter = "HTM-Dateien|*.htm|HTML-Dateien|*.html";
            oFile.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;
            oFile.Multiselect = true;
            Nullable<bool> result = oFile.ShowDialog();

            if (result == true)
            {
                int count = 1;
                foreach (string path in oFile.FileNames)
                {
                    ExcelFile ef = ExcelFile.Load(path);
                    string dateiname = path.Split('.')[0];
                    ef.Save(string.Format("{0}{1}.xlsx", dateiname, count));
                    count++;


                    hDoc.LoadHtml(path);
                    foreach (var table in hDoc.DocumentNode.SelectNodes("//table"))
                    {
                        foreach (var row in table.SelectNodes("tr"))
                        {
                            foreach (var cell in row.SelectNodes("td"))
                                MessageBox.Show(cell.InnerText);
                        }
                    }
                }
            }
        }
    }
}
