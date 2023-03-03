using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;


namespace Excel
{
    internal class Product
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

        public string ShortDescription => ReadShortDecription();

        private string ReadShortDecription()
        {
            HtmlNodeCollection description = html.DocumentNode.SelectNodes("//div[contains(@class, 'col-12 descriptiondiv')]");

            return Validator.InnerTextControl(description);
        }

        public string Model => ReadModel();

        private string ReadModel()
        {
            HtmlNodeCollection model = html.DocumentNode.SelectNodes("//span[contains(@data-io, 'model')]");

            return Validator.InnerTextControl(model);

        }

        public string Price => ReadPrice();

        private string ReadPrice()
        {
            HtmlNodeCollection price = html.DocumentNode.SelectNodes("//div[contains(@class, 'product-price')]");
            return Validator.InnerTextControl(price);
        }

        public string Weight => ReadWeight();

        private string ReadWeight()
        {
            HtmlNodeCollection weight = html.DocumentNode.SelectNodes("//li[contains(@class, 'product-weight')]");
            
            return Validator.InnerTextControl(weight);
        }

        public string Brand => ReadBrand();

        private string ReadBrand()
        {
            HtmlNodeCollection brand = html.DocumentNode.SelectNodes("//div[contains(@class, 'brand-image product-manufacturer')]");

            return Validator.InnerTextControl(brand);
        }

        public string MinOrder => ReadMinOrder();

        private string ReadMinOrder()
        {
            HtmlNodeCollection minOrder = html.DocumentNode.SelectNodes("//div[contains(@class, 'minimum alert alert-info')]");

            return Validator.InnerTextControl(minOrder);
        }


        public string Picture => ReadPicture();

        private string ReadPicture()
        {
            HtmlNodeCollection picture = html.DocumentNode.SelectNodes("//div[contains(@class, 'swiper-slide')]//img");

            return Validator.PictureControl(picture);

        }

        public string Title => ReadTitle();

        private string ReadTitle()
        {
            HtmlNodeCollection node = html.DocumentNode.SelectNodes("h1");

            return Validator.InnerTextControl(node);
        }

        public string Barcode => ReadBarCode();

        private string ReadBarCode()
        {
            HtmlNodeCollection barCode = html.DocumentNode.SelectNodes("//span[contains(@data-io, 'sku')]");

            return Validator.InnerTextControl(barCode);

            
        }
    }
}