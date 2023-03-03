using HtmlAgilityPack;
using System;
using System.Linq;

namespace Excel
{
    internal class Validator
    {
        internal static string InnerTextControl(HtmlNodeCollection description)
        {

            if (description != null)
            {
                return description.First().InnerText;

            }
            return string.Empty;
        }


        internal static string PictureControl(HtmlNodeCollection picture)
        {

            if (picture != null)
            {
                return picture.First().GetAttributeValue("src", " ");

            }
            return string.Empty;
        }
    }
}