using HtmlAgilityPack;
using System;
using System.Text.RegularExpressions;
using Excel1 = Microsoft.Office.Interop.Excel;

namespace Promotions
{
    class Program
    {
        static void Main(string[] args)
        {

            Excel1.Application xlApp = new Excel1.Application();

            Excel1.Workbook xlWorkBook;
            Excel1.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel1.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\Svetlio-PC\Documents\Book1.xlsx", 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel1.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlApp.Visible = true;
            Excel1.Range xlRange = xlWorkSheet.UsedRange;


            HtmlWeb web = new HtmlWeb();

            var htmlDoc = web.Load("https://tashev-galving.com/promos/sola");

            // var nodes = htmlDoc.DocumentNode.SelectNodes("//div[@class='card-content']//del[@class='has-text-grey-light is-pulled-left']");

            var nodes = htmlDoc.DocumentNode.SelectNodes("//div[@class='card-content']");
            if (nodes != null)
            {
                //Console.WriteLine(nodes.Count);
                // string nalichnost = node.InnerText.Substring(0, node.InnerText.IndexOf("(") + 1).Trim();

                // string resultString = Regex.Match(nalichnost, @"\d+").Value;

                foreach (var item in nodes)
                {
                    string url = item.ChildNodes[1].InnerHtml;

                    string price = item.ChildNodes[3].InnerText;

                    string dicountPrice = item.ChildNodes[5].InnerText;
                    dicountPrice = Regex.Match(dicountPrice, @"\d+").Value;

                    Console.WriteLine(url);


                    // Console.WriteLine(item.InnerHtml);
                }
            }
        }
    }
}
