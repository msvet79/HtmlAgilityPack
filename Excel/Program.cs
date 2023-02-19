using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using HtmlAgilityPack;
using Excel1 = Microsoft.Office.Interop.Excel;

namespace Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string url = "https://www.valerii.com/motofreza-benz-mini2-2kw52cc-42659";
            // HtmlWeb web = new HtmlWeb();
            // HtmlDocument html = web.Load(url);
            //regular - price
            //  HtmlNodeCollection titleNodes = html.DocumentNode.SelectNodes("//span[contains(@data-io, 'sku')]");
            //Console.WriteLine(html.ParsedText);
            // HtmlNodeCollection description = html.DocumentNode.SelectNodes("//div[contains(@class, 'col-12 descriptiondiv')]");

            // HtmlNodeCollection model = html.DocumentNode.SelectNodes("//span[contains(@data-io, 'model')]");

            // HtmlNodeCollection weight = html.DocumentNode.SelectNodes("//li[contains(@class, 'product-weight')]");

            // HtmlNodeCollection price = html.DocumentNode.SelectNodes("//div[contains(@class, 'product-price')]");

            // HtmlNodeCollection brand = html.DocumentNode.SelectNodes("//div[contains(@class, 'brand-image product-manufacturer')]");


            // HtmlNodeCollection minOrder = html.DocumentNode.SelectNodes("//div[contains(@class, 'minimum alert alert-info')]");

            //foreach (HtmlNode node in brand)
            //{
            //Console.WriteLine(node.InnerText);
            //Console.WriteLine();
            //}

            // HtmlNode node = html.DocumentNode.SelectSingleNode("h1");
            //Console.WriteLine(titleNodes.FirstOrDefault().InnerText);
            //Console.WriteLine(node.OuterHtml);


            //Excel1.Application xlApp = new Excel1.Application();

            //Excel1.Workbook xlWorkBook;
           // Excel1.Worksheet xlWorkSheet;
           // object misValue = System.Reflection.Missing.Value;

          //  xlApp = new Excel1.Application();
           // xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\Svetlio-PC\Documents\Svet8.xlsx", 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            //xlWorkSheet = (Excel1.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //Excel1.Range xlRange = xlWorkSheet.UsedRange;


          //  for (int i = 488; i <= 724; i++)
          //  {

                //Console.Write("\r\n");

                //write the value to the console
                // if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null)
                //Console.Write(xlRange.Cells[i, 1].Value2.ToString() + "\t");
               // string url = xlRange.Cells[i, 2].Value;
                HtmlWeb web = new HtmlWeb();
                HtmlDocument html = web.Load(url);
            // HtmlNodeCollection description = html.DocumentNode.SelectNodes("//div[contains(@class, 'col-12 descriptiondiv')]");
            HtmlNodeCollection picture = html.DocumentNode.SelectNodes("//div[contains(@class, 'swiper-slide')]//img");

            Console.WriteLine(picture.First().GetAttributeValue("src"," "));
            

        }

          //  for (int i = 1; i < 16211; i++)
           // {
             //   try
             //   {
                 //   Product valerii = new Product(xlRange.Cells[i, 1].value);
                   // xlRange.Cells[i, 2] = valerii.ShortDescription;

                   // xlRange.Cells[i, 3] = valerii.DetailedDescription;

                  //  xlRange.Cells[i, 4] = valerii.Price;

                   // xlRange.Cells[i, 5] = valerii.TashevCode;

                  //  xlRange.Cells[i, 6] = valerii.Title;

                   // xlRange.Cells[i, 7] = valerii.ProducerCode;

                  //  xlRange.Cells[i, 8] = valerii.Producer;

                  //  xlRange.Cells[i, 9] = valerii.Picture;

                   // xlRange.Cells[i, 10] = valerii.Discount;
                  //  Console.WriteLine(i);
                   // Thread.Sleep(3000);



               // }
                //catch (Exception ex)
              //  {

                  //  Console.Write(ex.Message);
                //}

                //finally
              //  {
                    //xlWorkBook.Save();
                    // xlWorkBook.Close();

              //  }



           // }



        //}
    }
}

