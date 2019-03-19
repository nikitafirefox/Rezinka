using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using ProjectX;

namespace ProjectX.Dict
{
    public class Marking
    {

        public string Id { get; private set; }
        public string Width { get;private set; }
        public string Height { get; private set; }
        public string Diameter { get; private set; }
        public string SpeedIndex { get; set; }
        public string LoadIndex { get; set; }
        public string Country { get; set; }
        public string TractionIndex { get; set; }
        public string TemperatureIndex { get; set; }
        public string TreadwearIndex { get; set; }
        public bool ExtraLoad { get; set; }
        public bool RunFlat { get; set; }
        public string FlangeProtection { get; set; }

        private HashSet<string> Accomadations { get; set; }

        public Marking(XmlNode x)
        {
            Accomadations = new HashSet<string>();

            Id = x.Attributes.GetNamedItem("id").Value;
            Width = x.SelectSingleNode("width").InnerText;
            Height = x.SelectSingleNode("height").InnerText;
            Diameter = x.SelectSingleNode("diameter").InnerText;
            SpeedIndex = x.SelectSingleNode("speedIndex").InnerText;
            LoadIndex = x.SelectSingleNode("loadIndex").InnerText;
            Country = x.SelectSingleNode("country").InnerText;
            TractionIndex = x.SelectSingleNode("tractionIndex").InnerText;
            TemperatureIndex = x.SelectSingleNode("temperatureIndex").InnerText;
            TreadwearIndex = x.SelectSingleNode("treadwearIndex").InnerText;
            FlangeProtection = x.SelectSingleNode("flangeProtection").InnerText;

            ExtraLoad = bool.Parse(x.SelectSingleNode("extraLoad").Attributes.GetNamedItem("value").Value);
            RunFlat = bool.Parse(x.SelectSingleNode("runFlat").Attributes.GetNamedItem("value").Value);

            foreach (XmlNode xNode in x.SelectSingleNode("accomadations").ChildNodes)
            {
                Accomadations.Add(xNode.InnerText);
            }



        }

        public Marking(string id, string width, string height, string diameter, string speedIndex,
            string loadIndex, string country, string tractionIndex, string temperatureIndex, string treadwearIndex,
            bool extraLoad, bool runFlat, string flangeProtection){

            Accomadations = new HashSet<string>();

            Id = id;
            Width = width;
            Height = height;
            Diameter = diameter;
            SpeedIndex = speedIndex;
            LoadIndex = loadIndex;
            Country = country;
            TractionIndex = tractionIndex;
            TemperatureIndex = temperatureIndex;
            TreadwearIndex = treadwearIndex;
            ExtraLoad = extraLoad;
            RunFlat = runFlat;
            FlangeProtection = flangeProtection;
        }

        public override string ToString()
        {
            return "ID: " + Id + '\n'
                + "Marking: " + Width+"\\"+ Height+"R"+ Diameter + '\n'
                + "SpeedIndex: " + SpeedIndex + '\n'
                + "LoadIndex: " + LoadIndex + '\n'
                + "Country: " + Country + '\n'
                + "TractionIndex: " + TractionIndex + '\n'
                + "TemperatureIndex: " + TemperatureIndex + '\n' 
                + "TreadwearIndex: " + TreadwearIndex + '\n'
                + "ExtraLoad: " + ExtraLoad + '\n'
                + "RunFlat: " + RunFlat + '\n'
                + "FlangeProtection: " + FlangeProtection + '\n'
                + "Variations: \n\t" + String.Join("\n\t", Accomadations);

        }

        public XmlElement GetXmlNode(XmlDocument document) {
            XmlElement element = document.CreateElement("marking");
            element.SetAttribute("id", Id);

            XmlElement e;

            e = document.CreateElement("width");
            e.InnerText = Width;
            element.AppendChild(e);
            e = document.CreateElement("height");
            e.InnerText = Height;
            element.AppendChild(e);
            e = document.CreateElement("diameter");
            e.InnerText = Diameter;
            element.AppendChild(e);
            e = document.CreateElement("speedIndex");
            e.InnerText = SpeedIndex;
            element.AppendChild(e);
            e = document.CreateElement("loadIndex");
            e.InnerText = LoadIndex;
            element.AppendChild(e);
            e = document.CreateElement("country");
            e.InnerText = Country;
            element.AppendChild(e);
            e = document.CreateElement("tractionIndex");
            e.InnerText = TractionIndex;
            element.AppendChild(e);
            e = document.CreateElement("temperatureIndex");
            e.InnerText = TemperatureIndex;
            element.AppendChild(e);
            e = document.CreateElement("treadwearIndex");
            e.InnerText = TreadwearIndex;
            element.AppendChild(e);
            e = document.CreateElement("flangeProtection");
            e.InnerText = FlangeProtection;
            element.AppendChild(e);

            e = document.CreateElement("extraLoad");
            e.SetAttribute("value", ExtraLoad.ToString());
            element.AppendChild(e);
            e = document.CreateElement("runFlat");
            e.SetAttribute("value", RunFlat.ToString());
            element.AppendChild(e);

            XmlElement accomadationsElement;
            element.AppendChild(accomadationsElement = document.CreateElement("accomadations"));
            foreach (string item in Accomadations)
            {
                e = document.CreateElement("accomadation");
                e.InnerText = item;
                accomadationsElement.AppendChild(e);
            }

            return element;
        }

        public void AddStringValue(string value)
        {
            Accomadations.Add(value);
        }

        public void Set(string speedIndex,
            string loadIndex, string country, string tractionIndex, string temperatureIndex, string treadwearIndex,
            bool extraLoad, bool runFlat, string flangeProtection) {

            SpeedIndex = speedIndex;
            LoadIndex = loadIndex;
            Country = country;
            TractionIndex = tractionIndex;
            TemperatureIndex = temperatureIndex;
            TreadwearIndex = treadwearIndex;
            ExtraLoad = extraLoad;
            RunFlat = runFlat;
            FlangeProtection = flangeProtection;
        }

        public void DeleteStringValue(string value)
        {
            Accomadations.Remove(value);
        }


    }
}
