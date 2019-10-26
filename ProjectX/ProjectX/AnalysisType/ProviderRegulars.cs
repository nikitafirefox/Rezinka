using ProjectX.Dict;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace ProjectX.AnalysisType
{
    public class ProviderRegulars
    {
        private List<ProviderRegular> ProviderRegularsList { get; set; }
        private readonly string pathXML;

        public ProviderRegulars()
        {
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

            XmlNode xmlNode = xroot.GetElementsByTagName("providersRegulars").Item(0);
            foreach (XmlNode x in xmlNode.ChildNodes)
            {
                ProviderRegularsList.Add(new ProviderRegular(x));
            }
        }

        public void Save()
        {
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

        public string Add(string idProv)
        {
            if (ProviderRegularsList.FindIndex(x => x.IdProvier == idProv) == -1)
            {
                ProviderRegularsList.Add(new ProviderRegular(idProv));
            }

            return idProv;
        }

        public string Add(string idProv, string reg, int priority)
        {
            int index = ProviderRegularsList.FindIndex(x => x.IdProvier == idProv);
            ProviderRegular PR;
            if (index == -1)
            {
                ProviderRegularsList.Add(PR = new ProviderRegular(idProv));
            }
            else
            {
                PR = ProviderRegularsList[index];
            }

            string id = PR.Add(reg, priority);
            return id;
        }

        public string Add(string idProv, string idReg, string reg, int index, string name)
        {
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

        public string Add(string idProv, string idReg, string value, string name)
        {
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

            string id = PR.Add(idReg, value, name);
            return id;
        }

        public void AddPassString(string idProv, string passString)
        {
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

            PR.AddPassString(passString);
        }

        public ProviderRegular GetProviderRegularById(string id) {
            return ProviderRegularsList.Where(x => x.IdProvier == id).First();
        }

        public bool IsContainMarking(string idProv, string buffer)
        {
            var config = ProviderRegularsList.Where(x => x.IdProvier == idProv).First().GetListByPriority(0);

            config.Sort((x, y) => y.RegularString.Length.CompareTo(x.RegularString.Length));

            bool b = false;

            foreach (var item in config)
            {
                if (Regex.IsMatch(buffer, item.RegularString, RegexOptions.IgnoreCase))
                {
                    b = true;
                    break;
                }
            }

            return b;
        }

        public int CountMarking(string idProv, string buffer)
        {
            var config = ProviderRegularsList.Where(x => x.IdProvier == idProv).First().GetListByPriority(0);

            config.Sort((x, y) => y.RegularString.Length.CompareTo(x.RegularString.Length));

            int b = 0;

            foreach (var item in config)
            {
                if (Regex.IsMatch(buffer, item.RegularString, RegexOptions.IgnoreCase))
                {
                    b = Regex.Matches(buffer,item.RegularString, RegexOptions.IgnoreCase).Count;
                    break;
                }
            }

            return b;
        }

        public Dictionary<string, string> GetDictionary(string idProv, string buffer, out string outBuffer)
        {
            Dictionary<string, string> keyValues = new Dictionary<string, string>();
            outBuffer = buffer;

            var thisProviderRegular = ProviderRegularsList.Where(x => x.IdProvier == idProv).First();
            var priorityArray = thisProviderRegular.GetPrioritys().ToList();
            priorityArray.Sort();

            foreach (var item in priorityArray)
            {
                var config = ProviderRegularsList.Where(x => x.IdProvier == idProv).First().GetListByPriority(item);

                config.Sort((x, y) => y.RegularString.Length.CompareTo(x.RegularString.Length));

                PrimaryRegular primary = null;

                foreach (var itemConfig in config)
                {
                    if (Regex.IsMatch(outBuffer, itemConfig.RegularString, RegexOptions.IgnoreCase))
                    {
                        primary = itemConfig;
                        break;
                    }
                }

                if (primary == null) { continue; }

                string str = Regex.Match(outBuffer, primary.RegularString, RegexOptions.IgnoreCase).Value;
                outBuffer = outBuffer.Replace(str, "");

                keyValues = primary.GetDictionaryGroup(str, keyValues);

            }

            outBuffer = thisProviderRegular.Replace(outBuffer);

            return keyValues;
        }
    }

    public class ProviderRegular:IEnumerable
    {
        private GenId Gen { get; set; }
        public string IdProvier { get; private set; }

        private List<PrimaryRegular> PrimaryRegulars { get; set; }
        private List<string> PassString { get; set; }

        public ProviderRegular(XmlNode x)
        {
            PrimaryRegulars = new List<PrimaryRegular>();
            PassString = new List<string>();

            IdProvier = x.Attributes.GetNamedItem("id").Value;

            XmlNode xmlSet = x.SelectSingleNode("settings");
            Gen = new GenId(char.Parse(xmlSet.ChildNodes.Item(0).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(1).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(2).InnerText));

            foreach (XmlNode xNode in x.SelectSingleNode("regulars").ChildNodes)
            {
                PrimaryRegulars.Add(new PrimaryRegular(xNode));
            }

            foreach (XmlNode xNode in x.SelectSingleNode("passStrings").ChildNodes)
            {
                PassString.Add(xNode.InnerText);
            }
        }

        public ProviderRegular(string idProv)
        {
            IdProvier = idProv;
            Gen = new GenId('A', -1, 1);
            PrimaryRegulars = new List<PrimaryRegular>();
            PassString = new List<string>();
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

            XmlElement passStringsElem;
            element.AppendChild(passStringsElem = xmlDocument.CreateElement("passStrings"));
            XmlElement e;
            foreach (var item in PassString)
            {
                e = xmlDocument.CreateElement("passString");
                e.InnerText = item;
                passStringsElem.AppendChild(e);
            }

            return element;
        }

        public void AddPassString(string s)
        {
            
                PassString.Add(s);
        }

        public string AddPrimaryRegular(PrimaryRegular primary) {
            string id = Gen.NexVal();
            primary.Id = id;

            PrimaryRegulars.Add(primary);

            return id;

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

            if (i == -1)
            {
                throw new Exception();
            }

            PR = PrimaryRegulars[i];
            string id = PR.Add(reg, index, name);

            return id;
        }

        public string Add(string idReg, string value, string name)
        {
            int i = PrimaryRegulars.FindIndex(x => x.Id == idReg);
            PrimaryRegular PR;

            if (i == -1)
            {
                throw new Exception();
            }

            PR = PrimaryRegulars[i];
            string id = PR.Add(value, name);

            return id;
        }

        public List<PrimaryRegular> GetListByPriority(int v)
        {
            return PrimaryRegulars.Where(x => x.Priority == v).ToList();
        }

        public PrimaryRegular GetPrimaryRegularById(string id) {
            return PrimaryRegulars.Where(x => x.Id == id).First();
        }

        public int[] GetPrioritys()
        {
            return PrimaryRegulars.Select(x => x.Priority).Distinct().ToArray();
        }

        public void DeleteString(string s) {
            PassString.Remove(s);
        }

        public void ChangeString(string sNew, string sOld) {

            DeleteString(sOld);
            AddPassString(sNew);

        }

        public void DeletePrimaryRegular(string id) {
            PrimaryRegular primaryRegular = GetPrimaryRegularById(id);

            PrimaryRegulars.Remove(primaryRegular);
        }

        public string Replace(string str)
        {
            string outStr = str;
            foreach (var item in PassString)
            {
                string sov_item = Regex.Match(outStr, Regex.Escape(item), RegexOptions.IgnoreCase).Value;
                if (sov_item != "")
                {
                    outStr = (" " + outStr + " ").Replace(" " + sov_item + " ", " ").Trim();
                }
            }
            return outStr;
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)PrimaryRegulars).GetEnumerator();
        }

        public IEnumerable GetPassString() {
            return PassString;
        }
    }

    public abstract class AbstractRegular : IEnumerable
    {

        protected GenId GenForRegularParams { get; set; }

        public string Id { get;  set; }

        public string RegularString { get; set; }

       

        protected List<RegularParam> RegularParams { get; set; }

        public string Add(string reg, int index, string name)
        {
            string id = GenForRegularParams.NexVal();
            RegularParams.Add(new RegularParam(id, reg, index, name));
            return id;
        }

        public string Add(string value, string name)
        {
            string id = GenForRegularParams.NexVal();
            RegularParams.Add(new RegularParam(id, value, name));
            return id;
        }

        public string AddRegularParam(RegularParam regularParam) {

            string id = GenForRegularParams.NexVal();
            regularParam.Id = id;
            RegularParams.Add(regularParam);

            return id;
        }

        public void DeleteRegularParam(string id) {
            RegularParam param = RegularParams.Find(x => x.Id == id);
            RegularParams.Remove(param);
        }

        public RegularParam GetRegularParamById(string id) {
            return RegularParams.Where(x => x.Id == id).First();
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)RegularParams).GetEnumerator();
        }

        public Dictionary<string, string> GetDictionary(string str, Dictionary<string, string> dictionary) {

            Dictionary<string, string> keyValues = dictionary;

            foreach (RegularParam itemParam in RegularParams)
            {
                string val;




                if (itemParam.IsConstant)
                {
                    val = itemParam.Value;
                }
                else
                {
                    if (!Regex.IsMatch(str, itemParam.RegularString , RegexOptions.IgnoreCase))
                    {
                        continue;
                    }
                    val = Regex.Matches(str, itemParam.RegularString , RegexOptions.IgnoreCase)[itemParam.RegularIndex].Value;
                }

                if (keyValues.ContainsKey(itemParam.NameParam))
                {
                    keyValues[itemParam.NameParam] = val;
                }
                else
                {
                    keyValues.Add(itemParam.NameParam, val);
                }
            }

            return keyValues;
        }

        public abstract XmlNode GetXmlNode(XmlDocument xmlDocument);

    }

    public class PrimaryRegular: GroupRegular
    {
        
        public int Priority { get; set; }

        public PrimaryRegular(XmlNode x)
        {
            XmlNode xmlSet = x.SelectSingleNode("settingForRegular");
            GenForRegularParams = new GenId(char.Parse(xmlSet.ChildNodes.Item(0).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(1).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(2).InnerText));

            xmlSet = x.SelectSingleNode("settingsForGroup");
            GenIdForGroupRegulars = new GenId(char.Parse(xmlSet.ChildNodes.Item(0).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(1).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(2).InnerText));

            Id = x.Attributes.GetNamedItem("id").Value;

            Priority = int.Parse(x.Attributes.GetNamedItem("priority").Value);

            RegularString = x.Attributes.GetNamedItem("str").Value;

            RegularParams = new List<RegularParam>();
            GroupRegulars = new List<GroupRegular>();

            foreach (XmlNode xNode in x.SelectSingleNode("regularParams").ChildNodes)
            {
                RegularParams.Add(new RegularParam(xNode));
            }

            foreach (XmlNode xNode in x.SelectSingleNode("regularGroups").ChildNodes)
            {
                GroupRegulars.Add(new GroupRegular(xNode));
            }
        }

        public PrimaryRegular(string id, string reg, int priority)
        {
            Id = id;
            RegularString = reg;
            Priority = priority;
            GenForRegularParams = new GenId('A', -1, 1);
            RegularParams = new List<RegularParam>();
            GenIdForGroupRegulars = new GenId('A', -1, 1);
            GroupRegulars = new List<GroupRegular>();
        }

        public PrimaryRegular() {
            GenForRegularParams = new GenId('A', -1, 1);
            RegularParams = new List<RegularParam>();
            GenIdForGroupRegulars = new GenId('A', -1, 1);
            GroupRegulars = new List<GroupRegular>();
        }

        public override XmlNode GetXmlNode(XmlDocument xmlDocument)
        {
            XmlElement element = xmlDocument.CreateElement("regular");
            element.SetAttribute("id", Id);
            element.SetAttribute("priority", Priority.ToString());
            element.SetAttribute("str", RegularString);

            element.AppendChild(GenForRegularParams.GetXmlNode(xmlDocument, "settingForRegular"));
            element.AppendChild(GenIdForGroupRegulars.GetXmlNode(xmlDocument, "settingsForGroup"));


            XmlElement regularParamsElement;
            element.AppendChild(regularParamsElement = xmlDocument.CreateElement("regularParams"));
            foreach (var item in RegularParams)
            {
                regularParamsElement.AppendChild(item.GetXmlNode(xmlDocument));
            }

            XmlElement regularGroupsElement;
            element.AppendChild(regularGroupsElement = xmlDocument.CreateElement("regularGroups"));
            foreach (var item in GroupRegulars)
            {
                regularGroupsElement.AppendChild(item.GetXmlNode(xmlDocument));
            }

            return element;
        }

    }

    public class GroupRegular: AbstractRegular {

        protected GenId GenIdForGroupRegulars { get; set;}

        protected List<GroupRegular> GroupRegulars { get; set;}

        public GroupRegular() {

            GenForRegularParams = new GenId('A', -1, 1);
            RegularParams = new List<RegularParam>();
            GenIdForGroupRegulars = new GenId('A', -1, 1);
            GroupRegulars = new List<GroupRegular>();

        }

        public GroupRegular(string id, string reg)
        {
            Id = id;
            RegularString = reg;
            GenForRegularParams = new GenId('A', -1, 1);
            RegularParams = new List<RegularParam>();
            GenIdForGroupRegulars = new GenId('A', -1, 1);
            GroupRegulars = new List<GroupRegular>();
        }

        public GroupRegular(XmlNode x)
        {
            XmlNode xmlSet = x.SelectSingleNode("settingForRegular");
            GenForRegularParams = new GenId(char.Parse(xmlSet.ChildNodes.Item(0).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(1).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(2).InnerText));

            xmlSet = x.SelectSingleNode("settingsForGroup");
            GenIdForGroupRegulars = new GenId(char.Parse(xmlSet.ChildNodes.Item(0).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(1).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(2).InnerText));

            Id = x.Attributes.GetNamedItem("id").Value;

            RegularString = x.Attributes.GetNamedItem("str").Value;

            RegularParams = new List<RegularParam>();
            GroupRegulars = new List<GroupRegular>();

            foreach (XmlNode xNode in x.SelectSingleNode("regularParams").ChildNodes)
            {
                RegularParams.Add(new RegularParam(xNode));
            }

            foreach (XmlNode xNode in x.SelectSingleNode("regularGroups").ChildNodes)
            {
                GroupRegulars.Add(new GroupRegular(xNode));
            }

        }

        public string AddGroupRegular(string reg) {
            string id = GenIdForGroupRegulars.NexVal();
            GroupRegulars.Add(new GroupRegular(id, reg));
            return id;
        }

        public string AddGroupRegular(GroupRegular group) {
            string id = GenIdForGroupRegulars.NexVal();

            group.Id = id;

            GroupRegulars.Add(group);

            return id;
        }

        public GroupRegular GetGroupRegularById(string id) {
            return GroupRegulars.Where(x => x.Id == id).First();
        }

        public override XmlNode GetXmlNode(XmlDocument xmlDocument)
        {
            XmlElement element = xmlDocument.CreateElement("regular");
            element.SetAttribute("id", Id);
            element.SetAttribute("str", RegularString);

            element.AppendChild(GenForRegularParams.GetXmlNode(xmlDocument,"settingForRegular"));
            element.AppendChild(GenIdForGroupRegulars.GetXmlNode(xmlDocument,"settingsForGroup"));


            XmlElement regularParamsElement;
            element.AppendChild(regularParamsElement = xmlDocument.CreateElement("regularParams"));
            foreach (var item in RegularParams)
            {
                regularParamsElement.AppendChild(item.GetXmlNode(xmlDocument));
            }

            XmlElement regularGroupsElement;
            element.AppendChild(regularGroupsElement = xmlDocument.CreateElement("regularGroups"));
            foreach (var item in GroupRegulars) {
                regularGroupsElement.AppendChild(item.GetXmlNode(xmlDocument));
            }

            return element;
        }

        public Dictionary<string, string> GetDictionaryGroup(string str, Dictionary<string, string> dictionary)
        {
            Dictionary<string, string> keyValues = GetDictionary(str,dictionary);

            

            foreach (var item in GroupRegulars) {
                string str2 = Regex.Match(str, item.RegularString, RegexOptions.IgnoreCase).Value;
                keyValues = item.GetDictionaryGroup(str2,dictionary);
            }

            return dictionary;

        }

        public IEnumerable GetGroupRegular() {
            return GroupRegulars;
        }

        public void DeleteGroupRegular(string id) {
            GroupRegular group = GetGroupRegularById(id);
            GroupRegulars.Remove(group);

        }

    }

    public class RegularParam
    {
        public string Id { get; set; }
        public string RegularString { get; set; }
        public int RegularIndex { get; set; }
        public string NameParam { get; set; }
        public string Value { get; set; }
        public bool IsConstant { get; set; }

        public RegularParam() {}
   
        public RegularParam(XmlNode x)
        {
            Id = x.Attributes.GetNamedItem("id").Value;

            RegularString = x.SelectSingleNode("reg").InnerText;
            RegularIndex = int.Parse(x.SelectSingleNode("regIndex").InnerText);
            Value = x.SelectSingleNode("value").InnerText;
            IsConstant = bool.Parse(x.SelectSingleNode("isConstant").InnerText);
            NameParam = x.SelectSingleNode("name").InnerText;
        }

        public RegularParam(string id, string reg, int index, string name)
        {
            Id = id;
            RegularString = reg;
            RegularIndex = index;
            NameParam = name;
            IsConstant = false;
            Value = "";
        }

        public RegularParam(string id, string value, string name)
        {
            Id = id;
            RegularString = "";
            RegularIndex = 0;
            NameParam = name;
            IsConstant = true;
            Value = value;
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

            e = xmlDocument.CreateElement("isConstant");
            e.InnerText = IsConstant.ToString();
            element.AppendChild(e);

            e = xmlDocument.CreateElement("value");
            e.InnerText = Value;
            element.AppendChild(e);

            e = xmlDocument.CreateElement("name");
            e.InnerText = NameParam;
            element.AppendChild(e);

            return element;
        }
    }


}