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
using HtmlAgilityPack;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;

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
                using (Stream sr = File.Open(oFile.FileName, FileMode.Open))
                {
                    List<string> strList = new List<string>();
                    string str = string.Empty;
                    hDoc = new HtmlDocument();
                    hDoc.OptionReadEncoding = false;
                    hDoc.Load(sr, Encoding.UTF8);
                    foreach (var header in hDoc.DocumentNode.SelectNodes("//head"))
                    {
                        for (int i = header.ChildNodes.Count - 1; i >= 0; i--)
                        {
                            if (header.ChildNodes[i].InnerText == null || header.ChildNodes[i].InnerText == string.Empty)
                                header.ChildNodes[i].Remove();
                        }
                        using (StreamWriter sWrite = new StreamWriter(string.Format("{0}\\{1}", AppDomain.CurrentDomain.BaseDirectory, System.Guid.NewGuid().ToString())))
                        {
                            header.WriteTo(sWrite);
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
                var excelApp = new Excel.Application();

                excelApp.Visible = true;

                excelApp.Workbooks.Add();

                Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
                ArrayList columns = new ArrayList() { "B", "C", "D", "E", "F", "G", "H", "I" };

                var rowCounter = 1;


                foreach (string path in oFile.FileNames)
                {
                    using (Stream stream = File.Open(path, FileMode.Open))
                    {
                        hDoc.OptionReadEncoding = false;
                        hDoc.Load(stream, Encoding.UTF8);
                        foreach (var table in hDoc.DocumentNode.SelectNodes("//table"))
                        {
                            foreach (var row in table.SelectNodes("tr"))
                            {
                                foreach (var cell in row.SelectNodes("td"))
                                {
                                    var columnCounter = 0;
                                    if (cell.InnerHtml.Contains("table"))
                                    {

                                        foreach (var subTable in cell.SelectNodes("table"))
                                        {
                                            foreach (var subRow in subTable.SelectNodes("tr"))
                                            {
                                                foreach (var subCell in subRow.SelectNodes("td"))
                                                {
                                                    rowCounter++;
                                                    string[] strMessages = System.Net.WebUtility.HtmlDecode(subCell.InnerText).Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                                                    foreach (var str in strMessages)
                                                    {
                                                        if (string.IsNullOrWhiteSpace(str))
                                                            continue;
                                                        workSheet.Cells[rowCounter, columns[columnCounter]] = str;
                                                    }

                                                }

                                            }
                                        }
                                        columnCounter++;
                                    }
                                    else if (cell.InnerHtml.Equals(string.Empty))
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        /*string[] strMessages = System.Net.WebUtility.HtmlDecode(cell.InnerText).Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                                        foreach (var str in strMessages)
                                        {
                                            if (String.IsNullOrWhiteSpace(str))
                                                continue;

                                        }*/
                                    }

                                }
                            }
                        }
                    }
                }
            }
        }

    }
}
