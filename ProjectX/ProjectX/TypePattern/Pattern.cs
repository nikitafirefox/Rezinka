using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using ProjectX.Dict;

namespace ProjectX.TypePattern
{
    public class Patterns: IEnumerable{

        private GenId IdGen { get; set; }
        private readonly string pathXML;
        private List<Pattern> ListPatterns { get; set; }

        public Patterns() {
            ListPatterns = new List<Pattern>();
            pathXML = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\Patterns.xml");

            if (!File.Exists(pathXML))
            {
                string dataDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data");
                if (!Directory.Exists(dataDirectory)) { Directory.CreateDirectory(dataDirectory); }
                IdGen = new GenId('A', -1, 1);

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

        private void Init()
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(pathXML);
            XmlElement xroot = xmlDocument.DocumentElement;

            XmlNode xmlSet = xroot.SelectSingleNode("settings");
            IdGen = new GenId(Char.Parse(xmlSet.ChildNodes.Item(0).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(1).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(2).InnerText));

            XmlNode xmlNode = xroot.GetElementsByTagName("patterns").Item(0);
            foreach (XmlNode x in xmlNode.ChildNodes)
            {
                ListPatterns.Add(new Pattern(x));
            }
        }


        public void Save() {

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(pathXML);
            XmlElement xroot = xmlDocument.DocumentElement;
            xroot.RemoveAll();
            xroot.AppendChild(IdGen.GetXmlNode(xmlDocument));
            XmlElement providersElement;
            xroot.AppendChild(providersElement = xmlDocument.CreateElement("patterns"));
            foreach (var item in ListPatterns)
            {
                providersElement.AppendChild(item.GetXmlNode(xmlDocument));
            }
            xmlDocument.Save(pathXML);

        }

        public void Add(string head, string body) {
            ListPatterns.Add(new Pattern(IdGen.NexVal(), head, body));
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)ListPatterns).GetEnumerator();
        }

        public Pattern GetPattern(string id) {
            return ListPatterns.Find(x => x.Id == id);
        }
    }

    public class Pattern
    {
        public string Id { get; private set; }

        public string Head { get; set; }

        public string Body { get; set; }

        public Pattern(string id, string head, string body) {
            Id = id;
            Head = head;
            Body = body;
        }

        public Pattern(XmlNode x)
        {

            Id = x.Attributes.GetNamedItem("id").Value;
            Head = x.SelectSingleNode("head").InnerText;
            Body = x.SelectSingleNode("body").InnerText;

        }

        public XmlNode GetXmlNode(XmlDocument xmlDocument)
        {
            XmlElement element = xmlDocument.CreateElement("pattern");
            element.SetAttribute("id", Id);

            XmlElement e = xmlDocument.CreateElement("head");
            e.AppendChild(xmlDocument.CreateCDataSection(Head));
            element.AppendChild(e);
           
            e = xmlDocument.CreateElement("body");
            e.AppendChild(xmlDocument.CreateCDataSection(Body));
            element.AppendChild(e);

            return element;
        }
    }
}
