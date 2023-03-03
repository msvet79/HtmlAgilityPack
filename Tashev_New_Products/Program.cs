using System;
using System.Threading;
using HtmlAgilityPack;
using Excel1 = Microsoft.Office.Interop.Excel;

namespace Tashev_New_Products
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel1.Application xlApp = new Excel1.Application();
            string url = string.Empty;
            Excel1.Workbook xlWorkBook;
            Excel1.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel1.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\Svetlio-PC\Documents\New_Update_Missing.xlsx", 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel1.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlApp.Visible = true;
            Excel1.Range xlRange = xlWorkSheet.UsedRange;
            int counter = 0;

            for (int i = 1; i < 7; i++)
            {
                try
                {
                    url = (string)(xlRange.Cells[i, 4] as Excel1.Range).Value;
                    HtmlWeb web = new HtmlWeb();
                    HtmlDocument html = web.Load(url);
                    Thread.Sleep(20000);
                    HtmlNodeCollection nodes = html.DocumentNode.SelectNodes("//div[contains(@class, 'af-product-img-text')]//a[@href]");


                    for (int j = 0; j < nodes.Count; j++)
                    {
                        if (j % 2 != 0)
                        {
                            xlRange.Cells[j + 15 + counter, 1] = nodes[j].Attributes["href"].Value;
                        }
                    }
                    counter += nodes.Count;

                    Console.WriteLine(i);
                }
                catch (Exception ex)
                {

                    Console.WriteLine(ex.Message);
                    continue;
                }
               

            }


         
            
        }
    }
}
