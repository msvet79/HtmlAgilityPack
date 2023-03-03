using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using HtmlAgilityPack;
using Excel1 = Microsoft.Office.Interop.Excel;


namespace Tashev
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
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\Svetlio-PC\Documents\Updates.xlsx", 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel1.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlApp.Visible = true;
            Excel1.Range xlRange = xlWorkSheet.UsedRange;

            for (int i = 1; i < 44959; i++)
            {
                try
                {
                    Product tashev = new Product(xlRange.Cells[i, 1].value);
                  //xlRange.Cells[i, 2] = tashev.ShortDescription;

                 // xlRange.Cells[i, 3] = tashev.DetailedDescription;

                   xlRange.Cells[i, 4] = tashev.Price;

                   xlRange.Cells[i, 5] = tashev.TashevCode;

                  // xlRange.Cells[i, 6] = tashev.Title;

                   xlRange.Cells[i, 7] = tashev.ProducerCode;

                   xlRange.Cells[i, 8] = tashev.Producer;

                 // xlRange.Cells[i, 9] = tashev.Picture;

                 // xlRange.Cells[i, 10] = tashev.Discount;

                   xlRange.Cells[i, 11] = tashev.Quantity;
                    Console.WriteLine(i);
                    Thread.Sleep(1000);



                }
                catch (Exception ex)
                {

                    Console.Write(ex.Message);
                }

                finally
                {
                    //xlWorkBook.Save();
                   // xlWorkBook.Close();
                   
                }

                

            }



            xlWorkBook.Save();
           xlWorkBook.Close();





        }
    }
}
