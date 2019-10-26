using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.IO;
using System.Xml;

namespace ProjectX.TypePattern
{
    public static class AvitoGenerator
    {

        public static List<AvitoAd> Generate(List<Element> elements, GroupPattern groupPattern,
            int count, List<AvitoAd> avitoAds) {

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
                else if (int.Parse(item.Count) < 4) {
                    price = item.PriceForTwo;
                }

                var el = avitoAds.Find(x => x.IdProduct == item.IdProduct);
                if (el == null)
                {
                    avitoAds.Add(new AvitoAd()
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
                        Images = item.Images,
         


                    });

                    curCount++;
                }
                else if (el.Priority > item.Priority) {

                    avitoAds.Remove(el);

                    avitoAds.Add(new AvitoAd()
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
                        Images = item.Images,
         



                    });

                }

                if (curCount == count) {
                    break;
                }
            }

            return avitoAds;
        }

        public static string GetString(string str, Element element) {

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
            else if (element.Season != "") {
                str = str.Replace("<Season>", element.Season).Replace("<LowSeason>", element.Season.ToLower()).Replace("<UpSeason>", element.Season.First().ToString().ToUpper() + element.Season.Remove(0, 1))
                    .Replace("<Сезонность>", element.Season).Replace("<Сезонность(мал.)>", element.Season.ToLower()).Replace("<Сезоность(Бол.)>", element.Season.First().ToString().ToUpper() + element.Season.Remove(0, 1));
            }

            string com = element.Commercial == "Да" ? "C" : "";

            string seasWithSpikes = element.Season;

            string runFlat = element.ModelName.ToLower().Contains("runflat")? "" : "RunFlat";

            if (seasWithSpikes == "Зимние") {
                if (element.Spikes == "Да")
                {
                    seasWithSpikes = "Зимние шипованные";
                }
                else {
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
                .Replace("<ExtraLoad>", element.ExtraLoad.Replace("Да", "XL").Replace("Нет", "")).Replace("<XL>", element.ExtraLoad.Replace("Да","XL").Replace("Нет",""))
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
                .Replace("<LowSeasonWithSpikes>",seasWithSpikes.ToLower()).Replace("<Сезонность_с_шиповкой(мал)>", seasWithSpikes.ToLower())
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

        public static string AddGetString(string str) {
            str = str.Replace("<Жирный>", "<strong>").Replace("<Ж>", "<strong>").Replace("</Жирный>", "</strong>").Replace("</Ж>", "</strong>").Replace("<strong>", "<strong>").Replace("</strong>", "</strong>")
                .Replace("<Курсив>", "<em>").Replace("<К>", "<em>").Replace("</Курсив>", "</em>").Replace("</К>", "</em>").Replace("<K>", "<em>").Replace("</K>", "</em>").Replace("<em>", "<em>").Replace("</em>", "</em>")
                .Replace("<Маркированный список>", "<ul>").Replace("<MC>", "<ul>").Replace("<ul>", "<ul>").Replace("</Маркированный список>", "</ul>").Replace("</MC>", "</ul>").Replace("</ul>", "</ul>")
                .Replace("<Нумерованный список>", "<ol>").Replace("<НC>", "<ol>").Replace("<ol>", "<ol>").Replace("</Нумерованный список>", "</ol>").Replace("</НC>", "</ol>").Replace("</ol>", "</ol>")
                .Replace("<Элемемнт списка>", "<li>").Replace("<ЭС>", "<li>").Replace("<li>", "<li>").Replace("</Элемемнт списка>", "</li>").Replace("</ЭС>", "</li>").Replace("</li>", "</li>")
            ;

            

            return str;

        }

        public static string GetLogicString(string str, Element element) {
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

                if (Regex.Matches(val, "\\([^\\[\\]\\{\\}\\(\\)]*\\)").Count > 1) {
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
                    else {
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
                else {
                    str = str.Replace(val, "");
                }




            }

            return str;
        }

        public static void ToXML(string path, List<AvitoAd> avitoAds) {

            FileStream fs = new FileStream(path, FileMode.Create);
            XmlTextWriter xmlOut = new XmlTextWriter(fs, Encoding.Unicode)
            {
                Formatting = Formatting.Indented
            };

            xmlOut.WriteStartElement("Ads");
            xmlOut.WriteAttributeString("", "formatVersion", "", "3");
            xmlOut.WriteAttributeString("", "target", "", "Avito.ru");
            xmlOut.Close();
            fs.Close();

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(path);
            XmlElement xroot = xmlDocument.DocumentElement;

            AvitoConfig config = new AvitoConfig();

            foreach (var item in avitoAds)
            {
                XmlElement element = xmlDocument.CreateElement("Ad");

                XmlElement e = xmlDocument.CreateElement("Id");
                var s = item.Addition == "" ? "" : item.Addition.GetHashCode().ToString();
                e.InnerText = item.IdProduct.GetHashCode().ToString() + s;
                element.AppendChild(e);

                e = xmlDocument.CreateElement("AllowEmail");
                e.InnerText = "Да";
                element.AppendChild(e);

                e = xmlDocument.CreateElement("ManagerName");
                e.InnerText = "Сергей";
                element.AppendChild(e);

                e = xmlDocument.CreateElement("ContactPhone");
                e.InnerText = "+7 918 6355543";
                element.AppendChild(e);

                e = xmlDocument.CreateElement("Address");
                e.InnerText = "Россия, Краснодарский край, Краснодар, Сормовская, 7/C";
                element.AppendChild(e);

                e = xmlDocument.CreateElement("AdType");
                e.InnerText = "Товар приобретен на продажу";
                element.AppendChild(e);

                e = xmlDocument.CreateElement("Title");
                e.InnerText = item.Head;
                element.AppendChild(e);

                e = xmlDocument.CreateElement("Description");
                e.AppendChild(xmlDocument.CreateCDataSection(item.Body));
                element.AppendChild(e);

                e = xmlDocument.CreateElement("Category");
                e.InnerText = "Запчасти и аксессуары";
                element.AppendChild(e);

                e = xmlDocument.CreateElement("TypeId");
                e.InnerText = "10-048";
                element.AppendChild(e);

                e = xmlDocument.CreateElement("Price");
                e.InnerText = item.Price;
                element.AppendChild(e);

                e = xmlDocument.CreateElement("RimDiameter");
                e.InnerText = item.Diametr;
                element.AppendChild(e);

                var season = config.AvitoSeason(item.Seasson, item.Spikes == "Да");

                e = xmlDocument.CreateElement("TireType");
                e.InnerText = season;
                element.AppendChild(e);


                if (Regex.IsMatch(item.Width.Trim(), "^[0-9]{3}$")) {
                    e = xmlDocument.CreateElement("TireSectionWidth");
                    e.InnerText = item.Width;
                    element.AppendChild(e);
                }

                if (Regex.IsMatch(item.Height.Trim(), "^[0-9]{2,3}$"))
                {
                    e = xmlDocument.CreateElement("TireAspectRatio");
                    e.InnerText = item.Height;
                    element.AppendChild(e);
                }

                e = xmlDocument.CreateElement("Images");

                foreach (string str in item.Images) {
                    XmlElement el = xmlDocument.CreateElement("Image");
                    el.SetAttribute("url", str);
                    e.AppendChild(el);
                }

                foreach (string str in config.AddingImages)
                {
                    XmlElement el = xmlDocument.CreateElement("Image");
                    el.SetAttribute("url", str);
                    e.AppendChild(el);
                }

                element.AppendChild(e);

                xroot.AppendChild(element);

            }




            xmlDocument.Save(path);


        }

    }

    public class AvitoAd {

        public string IdProduct;
        public int Priority;
        public string Addition;


        public string Width;
        public string Height;
        public string Diametr;
        public string Spikes;
        public string Seasson;
        public string Price;

        public List<string> Images;

        public string Head { get; set; }
        public string Body { get; set; }

    }

    public class AvitoConfig{

        private readonly string pathXML;
        public List<AvitoConfigParam> AvitoConfigParams { get; set; }
        public List<string> AddingImages { get; set; }

        public AvitoConfig() {
            AvitoConfigParams = new List<AvitoConfigParam>();
            AddingImages = new List<string>();

            pathXML = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\AvitoConfig.xml");

            if (!File.Exists(pathXML))
            {
                string dataDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data");
                if (!Directory.Exists(dataDirectory)) { Directory.CreateDirectory(dataDirectory); }

                FileStream fs = new FileStream(pathXML, FileMode.Create);
                XmlTextWriter xmlOut = new XmlTextWriter(fs, Encoding.Unicode)
                {
                    Formatting = Formatting.Indented
                };
                xmlOut.WriteStartDocument();
                xmlOut.WriteStartElement("root");
                xmlOut.WriteEndElement();
                xmlOut.WriteEndDocument();
                xmlOut.Close();
                fs.Close();
                Save();
            }
            else
            {
                Init();
            }

        }

        public void Save() {

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(pathXML);
            XmlElement xroot = xmlDocument.DocumentElement;
            xroot.RemoveAll();
            XmlElement paramsElement;
            xroot.AppendChild(paramsElement = xmlDocument.CreateElement("params"));
            foreach (var item in AvitoConfigParams)
            {
                paramsElement.AppendChild(item.GetXmlNode(xmlDocument));
            }

            xroot.AppendChild(paramsElement = xmlDocument.CreateElement("imgs"));
            foreach (var item in AddingImages)
            {
                XmlElement e = xmlDocument.CreateElement("img");
                e.InnerText = item;
                paramsElement.AppendChild(e);
            }

            xmlDocument.Save(pathXML);
        }

        private void Init()
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(pathXML);
            XmlElement xroot = xmlDocument.DocumentElement;

            XmlNode xmlNode = xroot.GetElementsByTagName("params").Item(0);
            foreach (XmlNode x in xmlNode.ChildNodes)
            {
                AvitoConfigParams.Add(new AvitoConfigParam(x));
            }

            xmlNode = xroot.GetElementsByTagName("imgs").Item(0);
            foreach (XmlNode x in xmlNode.ChildNodes) {
                AddingImages.Add(x.InnerText);
            }
        }

        public string AvitoSeason(string season, bool spikes) {
            
            if (season == "") { return "NULL"; }
            var str = season;

            foreach (var item in AvitoConfigParams)
            {
                if ((spikes == item.Spikes) && (season == item.Season)) {
                    str = item.TotalSeason;
                    break;
                }
            }

            return str;
        }

        public void Add(string season, bool spikes, string total) {
            AvitoConfigParams.Add(new AvitoConfigParam(season,spikes,total));
        }

    }

    public class AvitoConfigParam {

        public bool Spikes { get; set;}
        public string Season { get; set;}
        public string TotalSeason { get; set;}

        public AvitoConfigParam(string season, bool spikes, string total) {
            Season = season;
            Spikes = spikes;
            TotalSeason = total;
        }

        public AvitoConfigParam(XmlNode x)
        {
            Spikes = bool.Parse(x.SelectSingleNode("Spikes").InnerText);
            Season = x.SelectSingleNode("Season").InnerText;
            TotalSeason = x.SelectSingleNode("Total").InnerText;
        }

        public XmlNode GetXmlNode(XmlDocument xmlDocument)
        {
            XmlElement element = xmlDocument.CreateElement("param");

            XmlElement e = xmlDocument.CreateElement("Spikes");
            e.InnerText = Spikes.ToString();
            element.AppendChild(e);

            e = xmlDocument.CreateElement("Season");
            e.InnerText = Season;
            element.AppendChild(e);

            e = xmlDocument.CreateElement("Total");
            e.InnerText = TotalSeason;
            element.AppendChild(e);

            return element;
        }
    }
}
