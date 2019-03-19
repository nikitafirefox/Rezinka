using ProjectX.Dict;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace ProjectX.AnalysisType
{



    public class ProviderRegulars
    {
        private List<ProviderRegular> ProviderRegularsList { get; set;}
        private readonly string pathXML;

        public ProviderRegulars() {

            pathXML = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\ProvidersRegulars.xml");
            ProviderRegularsList = new List<ProviderRegular>();

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
            else {
                Init();
            }
        }

        private void Init() {

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(pathXML);
            XmlElement xroot = xmlDocument.DocumentElement;

            XmlNode xmlNode = xroot.GetElementsByTagName("providersRegulars").Item(0);
            foreach (XmlNode x in xmlNode.ChildNodes)
            {
                ProviderRegularsList.Add(new ProviderRegular(x));
            }

        }

        public void Save() {

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(pathXML);
            XmlElement xroot = xmlDocument.DocumentElement;
            xroot.RemoveAll();

            XmlElement providersRegularsElement;
            xroot.AppendChild(providersRegularsElement = xmlDocument.CreateElement("providersRegulars"));
            foreach (var item in ProviderRegularsList)
            {
                providersRegularsElement.AppendChild(item.GetXmlNode(xmlDocument));
            }

            xmlDocument.Save(pathXML);

        }

        public string Add(string idProv) {

            if (ProviderRegularsList.FindIndex(x => x.IdProvier == idProv) == -1) {
                ProviderRegularsList.Add(new ProviderRegular(idProv));
            }

            return idProv;
        }

        public string Add(string idProv, string reg, int priority) {
            int index = ProviderRegularsList.FindIndex(x => x.IdProvier == idProv);
            ProviderRegular PR;
            if (index == -1)
            {
                ProviderRegularsList.Add(PR = new ProviderRegular(idProv));
            }
            else {
                PR = ProviderRegularsList[index];
            }

            string id = PR.Add(reg, priority);
            return  id;
        }

        public string Add(string idProv, string idReg,string reg, int index, string name) {
            int i = ProviderRegularsList.FindIndex(x => x.IdProvier == idProv);
            ProviderRegular PR;
            if (i == -1)
            {
                ProviderRegularsList.Add(PR = new ProviderRegular(idProv));
            }
            else
            {
                PR = ProviderRegularsList[i];
            }

            string id = PR.Add(idReg, reg, index, name);
            return id;
        }

        public bool IsContainMarking(string idProv, string buffer) {

            var config = ProviderRegularsList.Where(x => x.IdProvier == idProv).First().GetListByPriority(0);

            config.Sort((x, y) => x.RegularString.CompareTo(y.RegularString));

            config.Reverse();

            bool b = false;

            foreach (var item in config)
            {
                if (Regex.IsMatch(buffer, item.RegularString, RegexOptions.IgnoreCase)) {
                    b = true;
                    break;
                }
            }
          
            return b;
        }

        public Dictionary<string,string> GetDictionary(string idProv, string buffer,out string outBuffer) {
            Dictionary<string, string> keyValues = new Dictionary<string, string>();
            outBuffer = buffer;

            var priorityArray = ProviderRegularsList.Where(x => x.IdProvier == idProv).First().GetPrioritys();
            foreach (var item in priorityArray)
            {
                var config = ProviderRegularsList.Where(x => x.IdProvier == idProv).First().GetListByPriority(item);

                config.Sort((x, y) => x.RegularString.CompareTo(y.RegularString));

                config.Reverse();


                PrimaryRegular primary = null;

                foreach (var itemConfig in config)
                {
                    if (Regex.IsMatch(outBuffer, itemConfig.RegularString, RegexOptions.IgnoreCase))
                    {
                        primary = itemConfig;
                        break;
                    }
                }

                if (primary == null) { continue;}
                string str = Regex.Match(outBuffer, primary.RegularString, RegexOptions.IgnoreCase).Value;
                outBuffer.Replace(str, "");

                foreach (RegularParam itemParam in primary)
                {
                    string val = Regex.Matches(str, itemParam.RegularString, RegexOptions.IgnoreCase)[itemParam.RegularIndex].Value;
                    if (keyValues.ContainsKey(itemParam.NameParam))
                    {
                        keyValues[itemParam.NameParam] = val;
                    }
                    else {
                        keyValues.Add(itemParam.NameParam, val);
                    }
                     
                }

            }
            return keyValues;
        }
    }

    public class ProviderRegular
    {

        private GenId Gen { get; set; }
        public string IdProvier { get; private set; }

        private List<PrimaryRegular> PrimaryRegulars { get; set; }



        public ProviderRegular(XmlNode x)
        {
            PrimaryRegulars = new List<PrimaryRegular>();

            IdProvier = x.Attributes.GetNamedItem("id").Value;

            XmlNode xmlSet = x.SelectSingleNode("settings");
            Gen = new GenId(char.Parse(xmlSet.ChildNodes.Item(0).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(1).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(2).InnerText));

            foreach (XmlNode xNode in x.SelectSingleNode("regulars").ChildNodes)
            {
                PrimaryRegulars.Add(new PrimaryRegular(xNode));
            }
        }

        public ProviderRegular(string idProv)
        {
            IdProvier = idProv;
            Gen = new GenId('A', -1, 1);
            PrimaryRegulars = new List<PrimaryRegular>();
        }

        public XmlNode GetXmlNode(XmlDocument xmlDocument)
        {
            XmlElement element = xmlDocument.CreateElement("providerRegular");
            element.SetAttribute("id", IdProvier);

            element.AppendChild(Gen.GetXmlNode(xmlDocument));

            XmlElement regularsElement;
            element.AppendChild(regularsElement = xmlDocument.CreateElement("regulars"));
            foreach (var item in PrimaryRegulars)
            {
                regularsElement.AppendChild(item.GetXmlNode(xmlDocument));
            }

            return element;
        }

        public string Add(string reg, int priority)
        {
            string id = Gen.NexVal();
            PrimaryRegulars.Add(new PrimaryRegular(id, reg, priority));
            return id;
        }

        public string Add(string idReg, string reg, int index, string name)
        {
            int i = PrimaryRegulars.FindIndex(x => x.Id == idReg);
            PrimaryRegular PR;

            if (i == -1) {
                throw new Exception();
            }

            PR = PrimaryRegulars[i];
            string id = PR.Add(reg, index, name);

            return id;

        }

        public List<PrimaryRegular> GetListByPriority(int v)
        {
            return PrimaryRegulars.Where(x => x.Priority == v).ToList();
        }

        public int[] GetPrioritys()
        {
            return PrimaryRegulars.Select(x => x.Priority).Distinct().ToArray();
        }
    }

     public class PrimaryRegular : IEnumerable {

        private GenId Gen { get; set;}
        public string Id { get; private set;}

        public string RegularString { get; set;}
        public int Priority { get; set;}

        private List<RegularParam> RegularParams { get; set;}



        public PrimaryRegular(XmlNode x)
        {

            XmlNode xmlSet = x.SelectSingleNode("settings");
            Gen = new GenId(char.Parse(xmlSet.ChildNodes.Item(0).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(1).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(2).InnerText));

            Id = x.Attributes.GetNamedItem("id").Value;

            Priority = int.Parse(x.Attributes.GetNamedItem("priority").Value);

            RegularString = x.Attributes.GetNamedItem("str").Value;

            RegularParams = new List<RegularParam>();

            foreach (XmlNode xNode in x.SelectSingleNode("regularParams").ChildNodes)
            {
                RegularParams.Add(new RegularParam(xNode));
            }


        }

        public PrimaryRegular(string id, string reg, int priority)
        {
            Id = id;
            RegularString = reg;
            Priority = priority;
            Gen = new GenId('A', -1, 1);
            RegularParams = new List<RegularParam>();
        }

        public XmlNode GetXmlNode(XmlDocument xmlDocument)
        {
            XmlElement element = xmlDocument.CreateElement("regular");
            element.SetAttribute("id", Id);
            element.SetAttribute("priority", Priority.ToString());
            element.SetAttribute("str", RegularString);

            element.AppendChild(Gen.GetXmlNode(xmlDocument));

            XmlElement regularParamsElement;
            element.AppendChild(regularParamsElement = xmlDocument.CreateElement("regularParams"));
            foreach (var item in RegularParams)
            {
                regularParamsElement.AppendChild(item.GetXmlNode(xmlDocument));
            }

            return element;
        }

        public string Add(string reg, int index, string name)
        {
            string id = Gen.NexVal();
            RegularParams.Add(new RegularParam(id, reg,index, name));
            return id;
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)RegularParams).GetEnumerator();
        }
    }

     public class RegularParam {


        public string Id { get; set; }
        public string RegularString { get; set;}
        public int RegularIndex { get; set; }
        public string NameParam { get; set; }

        public RegularParam(XmlNode x)
        {
            Id = x.Attributes.GetNamedItem("id").Value;

            RegularString = x.SelectSingleNode("reg").InnerText;
            RegularIndex = int.Parse(x.SelectSingleNode("regIndex").InnerText);
            NameParam = x.SelectSingleNode("name").InnerText;
        }

        public RegularParam(string id, string reg, int index, string name)
        {
            Id = id;
            RegularString = reg;
            RegularIndex = index;
            NameParam = name;
        }

        public XmlNode GetXmlNode(XmlDocument xmlDocument)
        {
            XmlElement element = xmlDocument.CreateElement("regularParam");
            element.SetAttribute("id", Id);

            XmlElement e = xmlDocument.CreateElement("reg");
            e.InnerText = RegularString;
            element.AppendChild(e);

            e = xmlDocument.CreateElement("regIndex");
            e.InnerText = RegularIndex.ToString();
            element.AppendChild(e);

            e = xmlDocument.CreateElement("name");
            e.InnerText = NameParam;
            element.AppendChild(e);

            return element;
        }
    }
}
