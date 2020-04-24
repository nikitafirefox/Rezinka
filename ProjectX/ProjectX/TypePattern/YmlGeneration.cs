using ProjectX.TireFitting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace ProjectX.TypePattern
{
    public class YmlGeneration
    {
        public static List<YmlAd> Generate(List<Element> elements, GroupPattern groupPattern,
            int count, List<YmlAd> ymlAds)
        {

            Patterns patterns = new Patterns();
            Pattern pattern;

            count = elements.Count > count ? count : elements.Count;
            int curCount = 0;

            foreach (var item in elements)
            {

                pattern = patterns.GetPattern(groupPattern.GetIdPatter());

                string price = item.Price;

                if (int.Parse(item.Count) == 1)
                {
                    price = item.PriceForOne;
                }
                else if (int.Parse(item.Count) < 4)
                {
                    price = item.PriceForTwo;
                }


                var el = ymlAds.Find(x => x.IdProduct == item.IdProduct);
                if (el == null)
                {
                    ymlAds.Add(new YmlAd()
                    {
                        IdProduct = item.IdProduct,
                        Diametr = item.Diameter,
                        Height = item.Height,
                        Seasson = item.Season,
                        Width = item.Width,
                        Priority = item.Priority,
                        Spikes = item.Spikes,
                        Head = GetLogicString(GetString(pattern.Head, item), item),
                        Body = GetLogicString(AddGetString(GetString(pattern.Body, item)), item),
                        Addition = item.Addition,
                        Price = price,
                        PriceForMrc = item.PriceForOne,
                        Images = item.Images,
                        IdProvider = item.ProvaiderId,
                        Brand = item.BrandName,
                        Time = item.TimeFromTo,
                        LoadIndex = item.LoadIndex,
                        SpeedIndex = item.SpeedIndex,
                        Count = item.Count,
                        FreeFitting = item.FreeFitting,

                    });

                    curCount++;
                }
                else if (el.Priority > item.Priority)
                {
                    ymlAds.Remove(el);

                    ymlAds.Add(new YmlAd()
                    {
                        IdProduct = item.IdProduct,
                        Diametr = item.Diameter,
                        Height = item.Height,
                        Seasson = item.Season,
                        Width = item.Width,
                        Priority = item.Priority,
                        Spikes = item.Spikes,
                        Head = GetLogicString(GetString(pattern.Head, item), item),
                        Body = GetLogicString(AddGetString(GetString(pattern.Body, item)), item),
                        Addition = item.Addition,
                        Price = price,
                        PriceForMrc = item.PriceForOne,
                        Images = item.Images,
                        IdProvider = item.ProvaiderId,
                        Brand = item.BrandName,
                        Time = item.TimeFromTo,
                        LoadIndex = item.LoadIndex,
                        SpeedIndex = item.SpeedIndex,
                        Count = item.Count,
                        FreeFitting = item.FreeFitting,

                    });

                }

                if (curCount == count)
                {
                    break;
                }
            }

            return ymlAds;
        }

        public static string GetString(string str, Element element)
        {

            str = str.Replace("<BrandDescription>", element.BrandDescription).Replace("<BDesc>", element.BrandDescription)
                .Replace("<Описание_производителя>", element.BrandDescription);
            str = str.Replace("<ModelDescription>", element.ModelDescription).Replace("<MDesc>", element.ModelDescription)
                .Replace("<Описание_модели>", element.ModelDescription);

            if (string.IsNullOrEmpty(element.Season))
            {
                str = str.Replace("<Season>", "").Replace("<LowSeason>", "").Replace("<UpSeason>", "")
                .Replace("<SeasonWihtSpikes>", "")
                .Replace("<LowSeasonWithSpikes>", "")
                .Replace("<UpSeasonWithSpikes>", "")
                .Replace("<Сезонность>", "").Replace("<Сезонность(мал.)>", "").Replace("<Сезоность(Бол.)>", "")
                .Replace("<Сезонность_с_шиповкой>", "")
                .Replace("<Сезонность_с_шиповкой(мал)>", "")
                .Replace("<Сезонность_с_шиповкой(Бол)>", "");
            }
            else if (element.Season != "")
            {
                str = str.Replace("<Season>", element.Season).Replace("<LowSeason>", element.Season.ToLower()).Replace("<UpSeason>", element.Season.First().ToString().ToUpper() + element.Season.Remove(0, 1))
                    .Replace("<Сезонность>", element.Season).Replace("<Сезонность(мал.)>", element.Season.ToLower()).Replace("<Сезоность(Бол.)>", element.Season.First().ToString().ToUpper() + element.Season.Remove(0, 1));
            }

            string com = element.Commercial == "Да" ? "C" : "";

            string seasWithSpikes = element.Season;

            string runFlat = element.ModelName.ToLower().Contains("runflat") ? "" : "RunFlat";

            if (element.RunFlat == "Нет")
            {

                if (element.ModelName.ToLower().Contains("runflat"))
                {
                    string s = Regex.Match(element.ModelName, "runflat", RegexOptions.IgnoreCase).Value;
                    element.ModelName = element.ModelName.Replace(s, "");
                }
            }

            if (seasWithSpikes == "Зимние")
            {
                if (element.Spikes == "Да")
                {
                    seasWithSpikes = "Зимние шипованные";
                }
                else
                {
                    seasWithSpikes = "Зимние нешипованные";
                }
            }


            str = str.Replace("<Accomadation>", element.Accomadation).Replace("<Ac>", element.Accomadation).Replace("<Амологация>", element.Accomadation)
                .Replace("<Addition>", element.Addition).Replace("<Add>", element.Addition).Replace("<Примичание>", element.Addition)
                .Replace("<BrandCountry>", element.BrandCountry).Replace("<BCountry>", element.BrandCountry).Replace("<Страна_производителя>", element.BrandCountry)
                .Replace("<BrandName>", element.BrandName).Replace("<BName>", element.BrandName).Replace("<Производитель>", element.BrandName)
                .Replace("<Commercial>", com).Replace("<C>", com).Replace("<Тип_шины(C)>", com)
                .Replace("<Count>", element.Count).Replace("<Ост>", element.Count).Replace("<Остаток>", element.Count)
                .Replace("<Date>", element.Date).Replace("<Дата>", element.Date)
                .Replace("<Diameter>", element.Diameter).Replace("<D>", element.Diameter).Replace("<d>", element.Diameter).Replace("<Диаметр>", element.Diameter).Replace("<Д>", element.Diameter).Replace("<д>", element.Diameter)
                .Replace("<ExtraLoad>", element.ExtraLoad.Replace("Да", "XL").Replace("Нет", "")).Replace("<XL>", element.ExtraLoad.Replace("Да", "XL").Replace("Нет", ""))
                .Replace("<xl>", element.ExtraLoad.Replace("Да", "xl").Replace("Нет", "")).Replace("<Повышенная_нагрузка>", element.ExtraLoad.Replace("Да", "xl").Replace("Нет", ""))
                .Replace("<FlangeProtection>", element.FlangeProtection).Replace("<FP>", element.FlangeProtection).Replace("<Защита борта>", element.FlangeProtection)
                .Replace("<Height>", element.Height).Replace("<H>", element.Height).Replace("<h>", element.Height).Replace("<Высота>", element.Height).Replace("<В>", element.Height).Replace("<в>", element.Height)
                .Replace("<LoadIndex>", element.LoadIndex).Replace("<LI>", element.LoadIndex).Replace("<Индекс_нагрузки>", element.LoadIndex)
                .Replace("<MarkingCountry>", element.MarkingCountry).Replace("<MCountry>", element.MarkingCountry).Replace("<Страна_производства>", element.MarkingCountry)
                .Replace("<ModelName>", element.ModelName).Replace("<MName>", element.ModelName).Replace("<Модель>", element.ModelName)
                .Replace("<MudSnow>", element.MudSnow.Replace("Да", "M+S").Replace("Нет", "")).Replace("<M+S>", element.MudSnow.Replace("Да", "M+S").Replace("Нет", ""))
                .Replace("<Price>", element.Price).Replace("<$>", element.Price).Replace("<Цена>", element.Price)
                .Replace("<RunFlat>", element.RunFlat.Replace("Да", runFlat).Replace("Нет", "")).Replace("<RF>", element.RunFlat.Replace("Да", runFlat).Replace("Нет", ""))
                .Replace("<RunFlatName>", element.RunFlatName).Replace("<Название_технологии_RunFlat>", element.RunFlatName)
                .Replace("<SeasonWihtSpikes>", seasWithSpikes).Replace("<Сезонность_с_шиповкой>", seasWithSpikes)
                .Replace("<LowSeasonWithSpikes>", seasWithSpikes.ToLower()).Replace("<Сезонность_с_шиповкой(мал)>", seasWithSpikes.ToLower())
                .Replace("<UpSeasonWithSpikes>", seasWithSpikes.First().ToString().ToUpper() + seasWithSpikes.Remove(0, 1)).Replace("<Сезонность_с_шиповкой(Бол)>", seasWithSpikes.First().ToString().ToUpper() + seasWithSpikes.Remove(0, 1))

                .Replace("<SpeedIndex>", element.SpeedIndex).Replace("<SI>", element.SpeedIndex).Replace("<Индекс_скорости>", element.SpeedIndex)
                .Replace("<Spikes>", element.Spikes.Replace("Да", "шипованные").Replace("Нет", "нешипованные"))
                .Replace("<Шип>", element.Spikes.Replace("Да", "шипованные").Replace("Нет", "нешипованные"))
                .Replace("<UpSpikes>", element.Spikes.Replace("Да", "Шипованные").Replace("Нет", "Нешипованные"))
                .Replace("<Шип(Бол)>", element.Spikes.Replace("Да", "Шипованные").Replace("Нет", "Нешипованные"))


                .Replace("<TemperatureIndex>", element.TemperatureIndex).Replace("<Temperature>", element.TemperatureIndex).Replace("<Индекс_температуры>", element.TemperatureIndex)
                .Replace("<Time>", element.Time).Replace("<Время>", element.Time)
                .Replace("<TimeTransit>", element.TimeTransit).Replace("<Срок_доставки>", element.TimeTransit)
                .Replace("<TractionIndex>", element.TractionIndex).Replace("<Traction>", element.TractionIndex).Replace("<Индекс_сцепления>", element.TractionIndex)
                .Replace("<TreadwearIndex>", element.TreadwearIndex).Replace("<Treadwear>", element.TreadwearIndex).Replace("<Индекс_ходимости>", element.TreadwearIndex)
                .Replace("<Type>", element.Type).Replace("<Тип_шины>", element.Type)
                .Replace("<WhileLetters>", element.WhileLetters).Replace("<WL>", element.WhileLetters).Replace("<Рефленные буквы>", element.WhileLetters)
                .Replace("<Width>", element.Width).Replace("<W>", element.Width).Replace("<w>", element.Width).Replace("<Ширина>", element.Width).Replace("<ш>", element.Width).Replace("<Ш>", element.Width)
                ;




            return str;
        }

        public static string AddGetString(string str)
        {
            str = str.Replace("<Жирный>", "").Replace("<Ж>", "").Replace("</Жирный>", "").Replace("</Ж>", "").Replace("<strong>", "").Replace("</strong>", "")
                .Replace("<Курсив>", "").Replace("<К>", "").Replace("</Курсив>", "").Replace("</К>", "").Replace("<K>", "").Replace("</K>", "").Replace("<em>", "").Replace("</em>", "")
                .Replace("<Маркированный список>", "").Replace("<MC>", "").Replace("<ul>", "").Replace("</Маркированный список>", "").Replace("</MC>", "").Replace("</ul>", "")
                .Replace("<Нумерованный список>", "").Replace("<НC>", "").Replace("<ol>", "").Replace("</Нумерованный список>", "").Replace("</НC>", "").Replace("</ol>", "")
                .Replace("<Элемемнт списка>", "").Replace("<ЭС>", "").Replace("<li>", "").Replace("</Элемемнт списка>", "").Replace("</ЭС>", "").Replace("</li>", "")
                .Replace("<p>", "").Replace("</p>", "").Replace("<br>", "");
            ;



            return str;

        }

        public static string GetLogicString(string str, Element element)
        {
            MatchCollection matchCollection = Regex.Matches(str, "\\{\\[[^\\[\\]\\{\\}\\(\\)]*\\][!=]='[^\\[\\]\\{\\}\\(\\)]*'\\([^\\[\\]\\{\\}\\(\\)]*\\)(\\([^\\[\\]\\{\\}\\(\\)]*\\))?\\}");
            foreach (Match item in matchCollection)
            {
                string val = item.Value;

                string param = Regex.Match(val, "\\[[^\\[\\]\\{\\}\\(\\)]*\\]").Value;
                param = param.Remove(param.Length - 1, 1).Remove(0, 1);

                string logic = Regex.Match(val, "[!=]=").Value;

                string lparam = Regex.Match(val, "'[^\\[\\]\\{\\}\\(\\)]*'").Value;
                lparam = lparam.Remove(lparam.Length - 1, 1).Remove(0, 1);

                string value1 = Regex.Matches(val, "\\([^\\[\\]\\{\\}\\(\\)]*\\)")[0].Value;
                value1 = value1.Remove(value1.Length - 1, 1).Remove(0, 1);
                string value2 = "";

                if (Regex.Matches(val, "\\([^\\[\\]\\{\\}\\(\\)]*\\)").Count > 1)
                {
                    value2 = Regex.Matches(val, "\\([^\\[\\]\\{\\}\\(\\)]*\\)")[1].Value;
                    value2 = value2.Remove(value2.Length - 1, 1).Remove(0, 1);
                }



                string valparam = element.GetParam(param);

                lparam = lparam.Trim() == "@NULL" ? "" : lparam;

                if (logic == "==")
                {
                    if (valparam == lparam)
                    {
                        str = str.Replace(val, value1);
                    }
                    else
                    {
                        str = str.Replace(val, value2);
                    }
                }
                else if (logic == "!=")
                {
                    if (valparam != lparam)
                    {
                        str = str.Replace(val, value1);
                    }
                    else
                    {
                        str = str.Replace(val, value2);
                    }
                }
                else
                {
                    str = str.Replace(val, "");
                }




            }

            return str;
        }

        public static void ToYandexMarket(string path, List<YmlAd> avtoruAds,bool mrc)
        {

            FileStream fs = new FileStream(path, FileMode.Create);
            XmlTextWriter xmlOut = new XmlTextWriter(fs, Encoding.UTF8)
            {
                Formatting = Formatting.Indented
            };


            xmlOut.WriteStartElement("yml_catalog");
            xmlOut.Close();
            fs.Close();

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(path);

            XmlDeclaration xmlDeclaration = xmlDocument.CreateXmlDeclaration("1.0", "utf-8", null);

            XmlElement xroot = xmlDocument.DocumentElement;

            xmlDocument.InsertBefore(xmlDeclaration, xroot);
            string stringDateTime = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + " " + DateTime.Now.ToShortTimeString();
            xroot.SetAttribute("date", stringDateTime);

            XmlElement xmlShop = xmlDocument.CreateElement("shop");
            XmlElement xmlShopElement = xmlDocument.CreateElement("name");
            xmlShopElement.InnerText = "ReZina123.ru";
            xmlShop.AppendChild(xmlShopElement);

            xmlShopElement = xmlDocument.CreateElement("company");
            xmlShopElement.InnerText = "ИП Кочкина О.В.";
            xmlShop.AppendChild(xmlShopElement);

            xmlShopElement = xmlDocument.CreateElement("url");
            xmlShopElement.InnerText = "http://rezina123.ru";
            xmlShop.AppendChild(xmlShopElement);

            XmlElement currencies = xmlDocument.CreateElement("currencies");
            xmlShopElement = xmlDocument.CreateElement("currency");
            xmlShopElement.SetAttribute("id", "RUB");
            xmlShopElement.SetAttribute("rate", "1");
            currencies.AppendChild(xmlShopElement);
            xmlShop.AppendChild(currencies);

            XmlElement categories = xmlDocument.CreateElement("categories");
            xmlShopElement = xmlDocument.CreateElement("category");
            xmlShopElement.SetAttribute("id", "1");
            xmlShopElement.InnerText = "Шины";
            categories.AppendChild(xmlShopElement);
            xmlShop.AppendChild(categories);

            XmlElement xmlOffers = xmlDocument.CreateElement("offers");

            List<string> freeFittingsYmlIds = new List<string>(); 

            foreach (var item in avtoruAds)
            {
                if (item.Addition != "")
                {
                    continue;
                }

                string ymlid = "id" + item.IdProduct.GetHashCode().ToString().Replace("-", "m");
                if (item.FreeFitting) {
                    freeFittingsYmlIds.Add(ymlid);
                }

                XmlElement element = xmlDocument.CreateElement("offer");
                element.SetAttribute("id", ymlid);
                element.SetAttribute("bid","50");

                string head = item.Head;

                XmlElement e = xmlDocument.CreateElement("name");
                e.InnerText = head;
                element.AppendChild(e);

                e = xmlDocument.CreateElement("vendor");
                e.InnerText = item.Brand;
                element.AppendChild(e);

                e = xmlDocument.CreateElement("url");
                e.InnerText = "http://rezina123.ru/i" + item.IdProduct.GetHashCode().ToString().Replace("-", "m") + "/";
                element.AppendChild(e);

                e = xmlDocument.CreateElement("price");
                e.InnerText = item.Price;
                element.AppendChild(e);

                e = xmlDocument.CreateElement("currencyId");
                e.InnerText = "RUB";
                element.AppendChild(e);

                e = xmlDocument.CreateElement("categoryId");
                e.InnerText = "1";
                element.AppendChild(e);

                e = xmlDocument.CreateElement("picture");
                e.InnerText = item.Images.First();
                element.AppendChild(e);

                string desc = "";

                if (item.Time == 0) {
                    desc = "Товар в наличии в Краснодаре.";
                } else if (item.Time == 1) {
                    desc = "Доставка 1 рабочий день. Товар в наличии в Краснодаре.";
                }
                else {
                    desc = "Доставка "+ item.Time + " раб. дн.";
                }

                if (!mrc) {
                    desc += " Бесплатный шиномонтаж или шиномонтаж со скидкой.";
                }

                e = xmlDocument.CreateElement("description");
                e.AppendChild(xmlDocument.CreateCDataSection(desc));
                element.AppendChild(e);

                if ((int.Parse(item.Count) >= 2))
                {

                    e = xmlDocument.CreateElement("min-quantity");
                    int s = int.Parse(item.Count) >= 4? 4 : 2;
                    e.InnerText = s.ToString();
                    element.AppendChild(e);

                }

                xmlOffers.AppendChild(element);

            }

            xmlShop.AppendChild(xmlOffers);

            XmlElement xmlGifts = xmlDocument.CreateElement("gifts");

            XmlElement giftElement = xmlDocument.CreateElement("gift");
            giftElement.SetAttribute("id", "g1promo");

            XmlElement el = xmlDocument.CreateElement("name");
            el.InnerText = "Бесплатный шиномонтаж";
            giftElement.AppendChild(el);

            xmlGifts.AppendChild(giftElement);

            xmlShop.AppendChild(xmlGifts);


            XmlElement xmlPromos = xmlDocument.CreateElement("promos");

            XmlElement promoElement = xmlDocument.CreateElement("promo");
            promoElement.SetAttribute("id","p1");
            promoElement.SetAttribute("type", "gift with purchase");

            XmlElement purchaseElement = xmlDocument.CreateElement("purchase");

            XmlElement requiredQuantityElement = xmlDocument.CreateElement("required-quantity");
            requiredQuantityElement.InnerText = "4";
            purchaseElement.AppendChild(requiredQuantityElement);

            foreach (string id in freeFittingsYmlIds) {
                XmlElement product = xmlDocument.CreateElement("product");
                product.SetAttribute("offer-id", id);
                purchaseElement.AppendChild(product);
            }

            promoElement.AppendChild(purchaseElement);

            XmlElement promoGiftsElement = xmlDocument.CreateElement("promo-gifts");

            XmlElement promoGiftElement = xmlDocument.CreateElement("promo-gift");
            promoGiftElement.SetAttribute("offer-id", "g1promo");
            promoGiftsElement.AppendChild(promoGiftElement);

            promoElement.AppendChild(promoGiftsElement);

            xmlPromos.AppendChild(promoElement);

            xmlShop.AppendChild(xmlPromos);

            xroot.AppendChild(xmlShop);

            xmlDocument.Save(path);


        }

    }


    public class YmlAd
    {

        public string IdProduct;
        public string IdProvider;
        public int Priority;
        public string Addition;

        public string Width;
        public string Height;
        public string Diametr;
        public string LoadIndex;
        public string SpeedIndex;
        public string Spikes;
        public string Seasson;
        public string Price;
        public string PriceForMrc;
        public string Count;

        public string Brand;

        public bool FreeFitting;


        public List<string> Images;

        public string Head { get; set; }
        public string Body;

        public int Time;
    }
}
