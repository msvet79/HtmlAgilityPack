using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tashev
{
    public static class Validator
    {
        public static string InnerTextControl(HtmlNodeCollection producerCode)
        {

            if (producerCode != null)
            {
                return producerCode.First().InnerText;

            }
            return string.Empty;
        }
    }
}
