using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Tashev
{
    public class Product
    {
        
        private string url;
        private readonly HtmlWeb web;

        private HtmlDocument html;

        public Product(string url)
        {
            this.url = url;

            this.web = new HtmlWeb();

            this.html = web.Load(this.url);
        }

        public string ShortDescription => ReadShortDescription();

        private string ReadShortDescription()
        {
            return html.GetElementbyId("prodbtns1").InnerText.Trim();
        }


        public string DetailedDescription => ReadDdetailedDescription();

        private string ReadDdetailedDescription()
        {
            HtmlNodeCollection detailedDescription = html.DocumentNode.SelectNodes("//p[contains(@style, 'margin-top:1rem;line-height: 1.7;')]");//подробно описание

            return Validator.InnerTextControl(detailedDescription);
        }

        public string Price => ReadPrice();

        private string ReadPrice()
        {
            HtmlNodeCollection cena = html.DocumentNode.SelectNodes("//strong[contains(@class, 'has-text-orange is-size-3 has-text-weight-bold')]");

            return Validator.InnerTextControl(cena);
        }

        public string ProducerCode => ReadProducerCode();

        private string ReadProducerCode()
        {
            
            HtmlNodeCollection producerCode = html.DocumentNode.SelectNodes("//span[contains(@style, 'display:inline; white-space: nowrap;')]");//код на производител и код на Ташев

            if (producerCode != null)
            {
               return producerCode.Last().InnerText;

            }
            return string.Empty;
        }

        public string TashevCode => ReadTashevCode();

        private string ReadTashevCode()
        {

            HtmlNodeCollection tashevCode = html.DocumentNode.SelectNodes("//span[contains(@style, 'display:inline; white-space: nowrap;')]");//код на производител и код на Ташев

           return  Validator.InnerTextControl(tashevCode);
        }

        public string Title => ReadTitle();

        private string ReadTitle()
        {
            HtmlNodeCollection ime = html.DocumentNode.SelectNodes("//title");

            return Validator.InnerTextControl(ime);
        }

        public string Discount => ReadDiscount();

        private string ReadDiscount()
        {
            HtmlNodeCollection discount = html.DocumentNode.SelectNodes("//strong[contains(@class, 'has-text-primary is-size-5 has-text-weight-bold')]");
            
            if (discount != null)
            {
                return discount.Last().InnerText;

            }
            return string.Empty;
        }

        public string Picture => ReadPicture();
        private string ReadPicture()
        {
            HtmlNodeCollection pictureUrl = html.DocumentNode.SelectNodes("//meta[contains(@property, 'og:image')]");

            if (pictureUrl != null)
            {

                return pictureUrl.First().GetAttributeValue("content", "");

            }
            return string.Empty;
        }

        public string Producer => ReadProducer();

        private string ReadProducer()
        {
            IEnumerable<HtmlNode> brand = html.GetElementbyId("prodbtns1").Descendants()
                .Where(n => n.Attributes.Any(a => a.Value.Contains("brand")));

            return brand.First().InnerText;
        }

        public string Quantity => ReadQuantity();

        private string ReadQuantity()
        {
            HtmlNodeCollection quantity = html.DocumentNode.SelectNodes("//*[text()[contains(., 'Ямбол')]]");
            
            if (quantity != null)
            {

                string nalichnost = quantity.First().InnerText.Substring(0, quantity.First().InnerText.IndexOf("(") + 1).Trim();

                string resultString = Regex.Match(nalichnost, @"\d+").Value;

                return resultString;

            }
            return string.Empty;
        }
    }
}

