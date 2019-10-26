using System;
using System.Collections.Generic;
using System.Xml;

namespace ProjectX.Dict
{
    public class Marking
    {
        public string Id { get; private set; }
        public string Width { get;  set; }
        public string Height { get;  set; }
        public string Diameter { get;  set; }
        public string SpeedIndex { get; set; }
        public string LoadIndex { get; set; }
        public string Country { get; set; }
        public string TractionIndex { get; set; }
        public string TemperatureIndex { get; set; }
        public string TreadwearIndex { get; set; }
        public bool ExtraLoad { get; set; }
        public bool RunFlat { get; set; }
        public string FlangeProtection { get; set; }
        public string Accomadation { get; set; }
        public bool Spikes { get; set; }

        public Marking(XmlNode x)
        {
            

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
            Spikes = bool.Parse(x.SelectSingleNode("spikes").Attributes.GetNamedItem("value").Value);
            Accomadation = x.SelectSingleNode("accomadation").InnerText;


        }

        public Marking(string id, string width, string height, string diameter, string speedIndex,
            string loadIndex, string country, string tractionIndex, string temperatureIndex, string treadwearIndex,
            bool extraLoad, bool runFlat, string flangeProtection,string accomadation,bool spikes)
        {

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
            Accomadation = accomadation;
            Spikes = spikes;

        }

        public override string ToString()
        {
            return "ID: " + Id + '\n'
                + "Marking: " + Width + "\\" + Height + "R" + Diameter + '\n'
                + "SpeedIndex: " + SpeedIndex + '\n'
                + "LoadIndex: " + LoadIndex + '\n'
                + "Country: " + Country + '\n'
                + "TractionIndex: " + TractionIndex + '\n'
                + "TemperatureIndex: " + TemperatureIndex + '\n'
                + "TreadwearIndex: " + TreadwearIndex + '\n'
                + "ExtraLoad: " + ExtraLoad + '\n'
                + "RunFlat: " + RunFlat + '\n'
                + "FlangeProtection: " + FlangeProtection + '\n'
                + "Accomadation: " + Accomadation + '\n'
                +"Spikes: " + Spikes;
        }

        public XmlElement GetXmlNode(XmlDocument document)
        {
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

            e = document.CreateElement("spikes");
            e.SetAttribute("value", Spikes.ToString());
            element.AppendChild(e);

            e = document.CreateElement("accomadation");
            e.InnerText = Accomadation;
            element.AppendChild(e);




            return element;
        }



        public void Set(string speedIndex,
            string loadIndex, string country, string tractionIndex, string temperatureIndex, string treadwearIndex,
            bool extraLoad, bool runFlat, string flangeProtection,string accomadation, bool spikes)
        {
            SpeedIndex = speedIndex;
            LoadIndex = loadIndex;
            Country = country;
            TractionIndex = tractionIndex;
            TemperatureIndex = temperatureIndex;
            TreadwearIndex = treadwearIndex;
            ExtraLoad = extraLoad;
            RunFlat = runFlat;
            FlangeProtection = flangeProtection;
            Accomadation = accomadation;
            Spikes = spikes;
        }


    }
}