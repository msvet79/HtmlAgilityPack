using System;
using System.Threading;
using HtmlAgilityPack;
using Excel1 = Microsoft.Office.Interop.Excel;

namespace Check_Tashev
{
    class Program
    {
        static void Main(string[] args)
        {
           
            HtmlWeb web = new HtmlWeb();
           

           Excel1.Application xlApp = new Excel1.Application();
           
            Excel1.Workbook xlWorkBook;
            Excel1.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

           xlApp = new Excel1.Application();
           xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\Svetlio-PC\Documents\Check.xlsx", 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

           xlWorkSheet = (Excel1.Worksheet)xlWorkBook.Worksheets.get_Item(1);
           Excel1.Range xlRange = xlWorkSheet.UsedRange;

            

          

            for (int i = 1; i < 147; i++)
            {
                string cellValue = (string)(xlWorkSheet.Cells[i, 2] as Excel1.Range).Value;
                HtmlDocument html = web.Load(cellValue); 
                HtmlNode price = html.DocumentNode.SelectSingleNode("//span[contains(@class, 'prod-price')]");

                HtmlNode url = html.DocumentNode.SelectSingleNode("//div[contains(@class, 'product-name')]//a[@href]");

                if (price != null)
                {


                    (xlWorkSheet.Cells[i, 3] as Excel1.Range).Value = price.InnerText.Trim();
                    // Console.Write(i); Console.Write(price.InnerText.Trim()); Console.WriteLine();
                    Console.WriteLine(i);
                }


                if (url != null)
                {


                    (xlWorkSheet.Cells[i, 4] as Excel1.Range).Value = url.Attributes["href"].Value;
                    
                }





                Thread.Sleep(1000);

            }


            xlWorkBook.Save();
            xlWorkBook.Close();

        }
    }
}
