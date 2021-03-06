﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;

namespace ProjectX.Dict
{
    public class Model:IEnumerable
    {
        public Marking this[string id] {
            get {
                return Markings.Find(x => x.Id == id);
            }
        }

        public string Id { get; private set; }
        public string Type { get; set; }

        public string Name { get; set; }
        public string Season { get; set; }
        public string Description { get; set; }
        public bool Commercial { get; set; }
        public string WhileLetters { get; set; }
        public bool MudSnow { get; set; }
        public string DescUrl { get; set; }

        public bool FreeFitting { get; set; }
        public int FittingDiameter { get; set; }

        private GenId IdGen { get; set; }
        private List<string> Variations { get; set; }
        private List<Marking> Markings { get; set; }
        private List<string> Images { get; set; }
        private List<string> Tags { get; set; }

        public Model(XmlNode x)
        {
            Images = new List<string>();
            Variations = new List<string>();
            Markings = new List<Marking>();
            Tags = new List<string>();

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


            try
            {
                DescUrl = x.SelectSingleNode("shinaGuide").InnerText;
            }
            catch {
                DescUrl = "";
            }

            foreach (XmlNode xNode in x.SelectSingleNode("variations").ChildNodes)
            {
                Variations.Add(xNode.InnerText);
            }

            foreach (XmlNode xNode in x.SelectSingleNode("images").ChildNodes)
            {
                Images.Add(xNode.Attributes.GetNamedItem("src").Value);
            }

            XmlNode xmlTags = null;
            XmlNode xmlFreeFitting = null;
            XmlNode xmlFittingDiameter = null;

            try
            {
                xmlTags = x.SelectSingleNode("tags");
                xmlFreeFitting = x.SelectSingleNode("freeFitting");
                xmlFittingDiameter = x.SelectSingleNode("fittingDiameter");
            }
            catch { }

            if (xmlTags != null)
            {
                foreach (XmlNode xNode in xmlTags.ChildNodes)
                {
                    Tags.Add(xNode.InnerText);
                }
            }

            Tags.Sort();

            FreeFitting = xmlFreeFitting != null ? bool.Parse(xmlFreeFitting.InnerText) : false;
            FittingDiameter = xmlFittingDiameter != null ? int.Parse(xmlFittingDiameter.InnerText) : 0;

            foreach (XmlNode xNode in x.SelectSingleNode("markings").ChildNodes)
            {
                Markings.Add(new Marking(xNode));
            }

            Markings.Sort((x1, x2) => x1.Width.CompareTo(x2.Width));

        }

        public Model(string id, string type, string name, string season, string description, bool commercial,
            string whileLetters, bool mudSnow)
        {
            Images = new List<string>();
            IdGen = new GenId('A', -1, 1);
            Variations = new List<string>();
            Markings = new List<Marking>();
            Tags = new List<string>();

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

            e = document.CreateElement("shinaGuide");
            e.InnerText = DescUrl;
            element.AppendChild(e);

            e = document.CreateElement("freeFitting");
            e.InnerText = FreeFitting.ToString();
            element.AppendChild(e);

            e = document.CreateElement("fittingDiameter");
            e.InnerText = FittingDiameter.ToString();
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


            XmlElement tagsElement;

            element.AppendChild(tagsElement = document.CreateElement("tags"));
            foreach (string item in Tags)
            {
                e = document.CreateElement("tag");
                e.InnerText = item;
                tagsElement.AppendChild(e);
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

        public IEnumerable<string> GetImages()
        {
            return Images;
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

            Variations.Sort((x, y) => y.Length.CompareTo(x.Length));
            variation = "";
            bool b = false;
            foreach (var item in Variations)
            {
               
                

                string regexStr = wordSearching ? "\\b" + Regex.Escape(item) + "\\b" : Regex.Escape(item);

                
                
                if (item.Contains(@"\") || item.Contains("+") || item.Contains("*") || item.Contains("?") 
                    || item.Contains("|") || item.Contains(".") || item.Contains("$") || item.Contains("{") || item.Contains("[") ||
                    item.Contains("(") || item.Contains(")") || item.Contains("#") || item.Contains("@")) {

                    regexStr = Regex.Escape(item);



                    if (wordSearching) {

                        char first = item[0];
                        char last = item[item.Length - 1];


                        if (first != '\\'&& first != '+'  && first != '*' && first != '?'
                            && first != '|' && first != '.' && first != '$' && first != '{' && first != '[' &&
                            first != '(' && first != ')' && first != '#' && first != '@') {
                            regexStr = "\\b" + regexStr;
                        }

                        if (last != '\\' && last != '+' && last != '*' && last != '?'
                            && last != '|' && last != '.' && last != '$' && last != '{' && last != '[' &&
                            last != '(' && last != ')' && last != '#' && last != '@')
                        {
                            regexStr = regexStr + "\\b";
                        }

                    }

                }
                

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

        public void GetWidths(List<string> res)
        {
            foreach (Marking item in Markings)
            {
                res.Add(item.Width);
            }
        }

        public void GetHeights(List<string> res)
        {
            foreach (Marking item in Markings)
            {
                res.Add(item.Height);
            }
        }

        public void GetDiametrs(List<string> res)
        {
            foreach (Marking item in Markings)
            {
                res.Add(item.Diameter);
            }
        }

        public void GetAccomadation(List<string> res) {
            foreach (Marking item in Markings)
            {
                res.Add(item.Accomadation);
            }
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)Markings).GetEnumerator();
        }

        public IEnumerable<string> GetVariations() {
            return Variations;
        }

        public IEnumerable<string> GetTags() {
            return Tags;
        }

        public void AddTag(string tag)
        {
            Tags.Add(tag);
        }

        public void DeleteTag(string tag)
        {
            Tags.Remove(tag);
        }

        public void ChangeTag(string tOld, string tNew)
        {
            DeleteTag(tOld);
            AddTag(tNew);
        }

        public void ChangeStringValue(string sOld, string sNew) {
            DeleteStringValue(sOld);
            AddStringValue(sNew);
        }

        public void ChangeImage(string iOld, string iNew) {
            DeleteImage(iOld);
            AddImage(iNew);
        }
    }
}