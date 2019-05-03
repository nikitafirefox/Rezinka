using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;

namespace ProjectX.Dict
{
    public class Model
    {
        public string Id { get; private set; }
        public string Type { get; set; }

        public string Name { get; set; }
        public string Season { get; set; }
        public string Description { get; set; }
        public bool Commercial { get; set; }
        public string WhileLetters { get; set; }
        public bool MudSnow { get; set; }

        private GenId IdGen { get; set; }
        private HashSet<string> Variations { get; set; }
        private List<Marking> Markings { get; set; }
        private HashSet<string> Images { get; set; }

        public Model(XmlNode x)
        {
            Images = new HashSet<string>();
            Variations = new HashSet<string>();
            Markings = new List<Marking>();

            Id = x.Attributes.GetNamedItem("id").Value;
            Type = x.Attributes.GetNamedItem("type").Value;

            XmlNode xmlSet = x.SelectSingleNode("settings");
            IdGen = new GenId(char.Parse(xmlSet.ChildNodes.Item(0).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(1).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(2).InnerText));

            Name = x.SelectSingleNode("name").InnerText;
            Description = x.SelectSingleNode("description").InnerText;

            WhileLetters = x.SelectSingleNode("whileLetters").InnerText;
            Season = x.SelectSingleNode("season").InnerText;

            Commercial = bool.Parse(x.SelectSingleNode("commercial").Attributes.GetNamedItem("value").Value);
            MudSnow = bool.Parse(x.SelectSingleNode("mudSnow").Attributes.GetNamedItem("value").Value);

            foreach (XmlNode xNode in x.SelectSingleNode("variations").ChildNodes)
            {
                Variations.Add(xNode.InnerText);
            }

            foreach (XmlNode xNode in x.SelectSingleNode("images").ChildNodes)
            {
                Images.Add(xNode.Attributes.GetNamedItem("src").Value);
            }

            foreach (XmlNode xNode in x.SelectSingleNode("markings").ChildNodes)
            {
                Markings.Add(new Marking(xNode));
            }
        }

        public Model(string id, string type, string name, string season, string description, bool commercial,
            string whileLetters, bool mudSnow)
        {
            Images = new HashSet<string>();
            IdGen = new GenId('A', -1, 1);
            Variations = new HashSet<string>();
            Markings = new List<Marking>();

            Id = id;
            Type = type;
            Name = name;
            Season = season;
            Description = description;
            Commercial = commercial;
            WhileLetters = whileLetters;
            MudSnow = mudSnow;
        }

        public override string ToString()
        {
            return "ID: " + Id + '\n'
                + "Type: " + Type + '\n'
                + "Name: " + Name + '\n'
                + "Season: " + Season + '\n'
                + "Commercial: " + Commercial + '\n'
                + "MudSnow: " + MudSnow + '\n'
                + "WhileLetters: " + WhileLetters + '\n'
                + "Description: " + Description + '\n'
                + "Image: \n\t" + String.Join("\n\t", Images) + '\n'
                + "Variations: \n\t" + String.Join("\n\t", Variations);
        }

        public XmlElement GetXmlNode(XmlDocument document)
        {
            XmlElement element = document.CreateElement("model");
            element.SetAttribute("id", Id);
            element.SetAttribute("type", Type);

            element.AppendChild(IdGen.GetXmlNode(document));

            XmlElement e = document.CreateElement("name");
            e.InnerText = Name;
            element.AppendChild(e);
            e = document.CreateElement("season");
            e.InnerText = Season;
            element.AppendChild(e);
            e = document.CreateElement("whileLetters");
            e.InnerText = WhileLetters;
            element.AppendChild(e);
            e = document.CreateElement("commercial");
            e.SetAttribute("value", Commercial.ToString());
            element.AppendChild(e);
            e = document.CreateElement("mudSnow");
            e.SetAttribute("value", MudSnow.ToString());
            element.AppendChild(e);
            e = document.CreateElement("description");
            e.AppendChild(document.CreateCDataSection(Description));
            element.AppendChild(e);

            XmlElement variationsElement;

            element.AppendChild(variationsElement = document.CreateElement("variations"));
            foreach (string item in Variations)
            {
                e = document.CreateElement("variation");
                e.InnerText = item;
                variationsElement.AppendChild(e);
            }

            XmlElement imagesElement;

            element.AppendChild(imagesElement = document.CreateElement("images"));
            foreach (string item in Images)
            {
                e = document.CreateElement("image");
                e.SetAttribute("src", item);
                imagesElement.AppendChild(e);
            }

            XmlElement markingsElement;
            element.AppendChild(markingsElement = document.CreateElement("markings"));
            foreach (Marking item in Markings)
            {
                markingsElement.AppendChild(item.GetXmlNode(document));
            }

            return element;
        }

        public string Add(string width, string height, string diameter, string speedIndex,
            string loadIndex, string country, string tractionIndex, string temperatureIndex, string treadwearIndex,
            bool extraLoad, bool runFlat, string flangeProtection, string accomadation, bool spikes)
        {
            string id = IdGen.NexVal();
            Markings.Add(new Marking(id, width, height, diameter, speedIndex, loadIndex, country, tractionIndex,
                temperatureIndex, treadwearIndex, extraLoad, runFlat, flangeProtection, accomadation, spikes));
            return id;
        }

        public void AddStringValue(string value)
        {
            Variations.Add(value);
        }


        public Marking SearchMarking(string width, string height, string diameter,string loadIndex, string speedIndex,string accomadation,bool spikes,bool runFalt, out bool isContain)
        {
            isContain = false;
            Marking marking = null;
            foreach (var item in Markings)
            {
                if (width == item.Width 
                    && height == item.Height 
                    && diameter == item.Diameter 
                    && accomadation == item.Accomadation
                    && loadIndex == item.LoadIndex
                    && speedIndex == item.SpeedIndex
                    && spikes == item.Spikes
                    && runFalt == item.RunFlat)
                {
                    marking = item;
                    isContain = true;
                    break;
                }
            }
            return marking;
        }


        public bool IsMatched(string parsingBufer, bool wordSearching, out string variation)
        {
            variation = "";
            bool b = false;
            foreach (var item in Variations)
            {
                string regexStr = wordSearching ? "\\b" + Regex.Escape(item) + "\\b" : Regex.Escape(item);
                if (Regex.IsMatch(parsingBufer, regexStr, RegexOptions.IgnoreCase))
                {
                    b = true;
                    variation = item;
                    break;
                }
            }
            return b;
        }

        public void Set(string name, string type, string season, string description, bool commercial,
            string whileLetters, bool mudSnow)
        {
            Name = name;
            Season = season;
            Description = description;
            Commercial = commercial;
            WhileLetters = whileLetters;
            MudSnow = mudSnow;
            Type = type;
        }

        public void Set(string idMarking, string speedIndex,
            string loadIndex, string country, string tractionIndex, string temperatureIndex, string treadwearIndex,
            bool extraLoad, bool runFlat, string flangeProtection, string accomadation,bool spikes)
        {
            try
            {
                Markings.Find(x => x.Id == idMarking).Set(speedIndex, loadIndex, country, tractionIndex, temperatureIndex,
                    treadwearIndex, extraLoad, runFlat, flangeProtection,accomadation,spikes);
            }
            catch (ArgumentNullException)
            {
                throw new ArgumentNullException();
            }
        }

        public void Delete(string idMarking)
        {
            Markings.Remove(Markings.Find(x => x.Id == idMarking));
        }

        public void DeleteStringValue(string value)
        {
            Variations.Remove(value);
        }


        public void AddImage(string image)
        {
            Images.Add(image);
        }

        public void DeleteImage(string image)
        {
            Images.Remove(image);
        }

        public Dictionary<string, object> GetValuesById(string idMarking)
        {
            Dictionary<string, object> resault = new Dictionary<string, object>();
            try
            {
                Marking marking = Markings.Find(x => x.Id == idMarking);
                resault.Add("Marking_Width", marking.Width);
                resault.Add("Marking_Height", marking.Height);
                resault.Add("Marking_Diameter", marking.Diameter);
                resault.Add("Marking_Country", marking.Country);
                resault.Add("Marking_LoadIndex", marking.LoadIndex);
                resault.Add("Marking_SpeedIndex", marking.SpeedIndex);
                resault.Add("Marking_ExtraLoad", marking.ExtraLoad.ToString().Replace("True", "Да").Replace("False", "Нет"));
                resault.Add("Marking_FlangeProtection", marking.FlangeProtection);
                resault.Add("Marking_RunFlat", marking.RunFlat.ToString().Replace("True", "Да").Replace("False", "Нет"));
                resault.Add("Marking_TemperatureIndex", marking.TemperatureIndex);
                resault.Add("Marking_TractionIndex", marking.TractionIndex);
                resault.Add("Marking_TreadwearIndex", marking.TreadwearIndex);
                resault.Add("Marking_Accomadation", marking.Accomadation);
                resault.Add("Marking_Spikes", marking.Spikes.ToString().Replace("True", "Да").Replace("False", "Нет"));

            }
            catch (Exception e)
            {
                throw e;
            }
            return resault;
        }
    }
}