using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ProjectX.TypePattern
{
    public static class AvtoruGenarator
    {

        public static List<AvtoruAd> Generate(List<Element> elements, GroupPattern groupPattern,
           int count, List<AvtoruAd> avitoAds)
        {

            Patterns patterns = new Patterns();
            Pattern pattern;

            count = elements.Count > count ? count : elements.Count;
            int curCount = 0;

            foreach (var item in elements)
            {

                pattern = patterns.GetPattern(groupPattern.GetIdPatter());


                var el = avitoAds.Find(x => x.IdProduct == item.IdProduct);
                if (el == null)
                {
                    avitoAds.Add(new AvtoruAd()
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
                        Price = item.Price,
                        Images = item.Images,
                        IdProvider = item.ProvaiderId


                    });

                    curCount++;
                }
                else if (el.Priority < item.Priority)
                {


                    avitoAds.Add(new AvtoruAd()
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
                        Price = item.Price,
                        Images = item.Images,
                        IdProvider = item.ProvaiderId



                    });

                }

                if (curCount == count)
                {
                    break;
                }
            }

            return avitoAds;
        }


        public static string GetString(string str, Element element)
        {

            str = str.Replace("<BrandCountry>", element.BrandCountry).Replace("<BCountry>", element.BrandCountry);
            str = str.Replace("<ModelDescription>", element.ModelDescription).Replace("<MDesc>", element.ModelDescription);

            if (element.Season == "")
            {
                str = str.Replace("<Low_Season>", "").Replace("<Up_Season>", "");
            }

            str = str.Replace("<Accomadation>", element.Accomadation).Replace("<Ac>", element.Accomadation)
                .Replace("<Addition>", element.Addition).Replace("<Add>", element.Addition)
                .Replace("<BrandDescription>", element.BrandDescription).Replace("<BDesc>", element.BrandDescription)
                .Replace("<BrandName>", element.BrandName).Replace("<BName>", element.BrandName)
                .Replace("<Commercial>", "C").Replace("<C>", "C")
                .Replace("<Count>", element.Count).Replace("<Ост>", element.Count)
                .Replace("<Date>", element.Date)
                .Replace("<Diameter>", element.Diameter).Replace("<D>", element.Diameter).Replace("<d>", element.Diameter)
                .Replace("<ExtraLoad>", "XL").Replace("<XL>", "XL").Replace("<xl>", "xl")
                .Replace("<FlangeProtection>", element.FlangeProtection).Replace("<FP>", element.FlangeProtection)
                .Replace("<Height>", element.Height).Replace("<H>", element.Height).Replace("<h>", element.Height)
                .Replace("<LoadIndex>", element.LoadIndex).Replace("<LI>", element.LoadIndex)
                .Replace("<MarkingCountry>", element.MarkingCountry).Replace("<MCountry>", element.MarkingCountry)
                .Replace("<ModelName>", element.ModelName).Replace("<MName>", element.ModelName)
                .Replace("<MudSnow>", "M+S").Replace("<M+S>", "M+S")
                .Replace("<Price>", element.Price).Replace("<$>", element.Price)
                .Replace("<RunFlat>", element.RunFlat).Replace("<RF>", element.RunFlatName)
                .Replace("<RunFlatName>", element.RunFlatName)
                .Replace("<Season>", element.Season).Replace("<Low_Season>", element.Season.ToLower()).Replace("<Up_Season>", element.Season.First().ToString().ToUpper() + element.Season.Remove(0, 1))
                .Replace("<SpeedIndex>", element.SpeedIndex).Replace("<SI>", element.SpeedIndex)

                .Replace("<Spikes>", element.Spikes.Replace("Да", "шипованные").Replace("Нет", "нешипованные"))
                .Replace("<Шип>", element.Spikes.Replace("Да", "шипованные").Replace("Нет", "нешипованные"))
                .Replace("<Up_Spikes>", element.Spikes.Replace("Да", "Шипованные").Replace("Нет", "Нешипованные"))
                .Replace("<Up_Шип>", element.Spikes.Replace("Да", "Шипованные").Replace("Нет", "Нешипованные"))


                .Replace("<TemperatureIndex>", element.TemperatureIndex).Replace("<Temper>", element.TemperatureIndex)
                .Replace("<Time>", element.Time)
                .Replace("<TimeTransit>", element.TimeTransit).Replace("<Дост>", element.TimeTransit)
                .Replace("<TractionIndex>", element.TractionIndex).Replace("<Tract>", element.TractionIndex)
                .Replace("<TreadwearIndex>", element.TreadwearIndex).Replace("<Treadwear>", element.TreadwearIndex)
                .Replace("<Type>", element.Type)
                .Replace("<WhileLetters>", element.WhileLetters).Replace("<WL>", element.WhileLetters)
                .Replace("<Width>", element.Width).Replace("<W>", element.Width).Replace("<w>", element.Width)
                ;




            return str;
        }

        public static string AddGetString(string str)
        {
            str = str.Replace("<Жирный>", "").Replace("<Ж>", "").Replace("</Жирный>", "").Replace("</Ж>", "").Replace("<strong>", "").Replace("</strong>", "")
                .Replace("<Курсив>", "").Replace("<К>", "").Replace("</Курсив>", "").Replace("</К>", "").Replace("<K>", "").Replace("</K>", "").Replace("<em>", "").Replace("</em>", "")
                .Replace("<Маркированный список>", "<ul>").Replace("<MC>", "<ul>").Replace("<ul>", "<ul>").Replace("</Маркированный список>", "</ul>").Replace("</MC>", "</ul>").Replace("</ul>", "</ul>")
                .Replace("<Нумерованный список>", "<ol>").Replace("<НC>", "<ol>").Replace("<ol>", "<ol>").Replace("</Нумерованный список>", "<ol>").Replace("</НC>", "<ol>").Replace("</ol>", "<ol>")
                .Replace("<Элемемнт списка>", "<li>").Replace("<ЭС>", "<li>").Replace("<li>", "<li>").Replace("</Элемемнт списка>", "</li>").Replace("</ЭС>", "</li>").Replace("</li>", "</li>")
            ;

            str = str.Replace(Environment.NewLine, " ");

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

    }

    public class AvtoruAd
    {

        public string IdProduct;
        public string IdProvider;
        public int Priority;
        public string Addition;

        public string Width;
        public string Height;
        public string Diametr;
        public string Spikes;
        public string Seasson;
        public string Price;

        public List<string> Images;

        public string Head;
        public string Body;

    }
}
