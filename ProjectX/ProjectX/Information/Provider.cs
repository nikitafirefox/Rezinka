using ProjectX.Dict;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace ProjectX.Information
{

   
    public class Providers: IEnumerable
    {

        public Provider this[int i] {
            get {
                return ProvidersList[i];
            }
        }

        public Provider this[string id] {
            get {
                return ProvidersList.Find(x => x.Id == id);
            }
        }

        private GenId GenId { get; set; }
        private readonly string pathXML;

        private List<Provider> ProvidersList { get; set; }

        public Providers()
        {
            ProvidersList = new List<Provider>();
            pathXML = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\Providers.xml");

            if (!File.Exists(pathXML))
            {
                string dataDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data");
                if (!Directory.Exists(dataDirectory)) { Directory.CreateDirectory(dataDirectory); }
                GenId = new GenId('A', -1, 1);

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
                InitProviders();
            }
        }

        private void InitProviders()
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(pathXML);
            XmlElement xroot = xmlDocument.DocumentElement;

            XmlNode xmlSet = xroot.SelectSingleNode("settings");
            GenId = new GenId(Char.Parse(xmlSet.ChildNodes.Item(0).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(1).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(2).InnerText));

            XmlNode xmlNode = xroot.GetElementsByTagName("providers").Item(0);
            foreach (XmlNode x in xmlNode.ChildNodes)
            {
                ProvidersList.Add(new Provider(x));
            }

            ProvidersList.Sort((x1, x2) => x1.Name.CompareTo(x2.Name));
        }

        public Dictionary<string, object> GetValuesById(string idProv, string idStock)
        {
            Provider provider;
            Dictionary<string, object> result;
            try
            {
                provider = ProvidersList.Find(x => x.Id == idProv);
                result = provider.GetValuesById(idStock);
            }
            catch (Exception e)
            {
                throw e;
            }
            result.Add("Provider_Name", provider.Name);
            result.Add("Provider_Priority", provider.Priority);
            return result;
        }

        public string[] GetProvidersId()
        {
            return ProvidersList.Select(x => x.Id).ToArray();
        }

        public string AddProvider(string name, int priority)
        {
            string id = GenId.NexVal();
            ProvidersList.Add(new Provider(id, name, priority));
            return id;
        }

        public string AddStock(string idProv, string name, string time)
        {
            try
            {
                string id = ProvidersList.Find(x => x.Id == idProv).AddStock(name, time);
                return id;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public string[] GetStocksId(string idProv)
        {
            string[] res = null;
            try
            {
                res = ProvidersList.Find(x => x.Id == idProv).GetStocksId();
            }
            catch { }
            return res;
        }

        public void Save()
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(pathXML);
            XmlElement xroot = xmlDocument.DocumentElement;
            xroot.RemoveAll();
            xroot.AppendChild(GenId.GetXmlNode(xmlDocument));
            XmlElement providersElement;
            xroot.AppendChild(providersElement = xmlDocument.CreateElement("providers"));
            foreach (var item in ProvidersList)
            {
                providersElement.AppendChild(item.GetXmlNode(xmlDocument));
            }
            xmlDocument.Save(pathXML);
        }

        public Dictionary<string, object> GetValuesById(string idProv)
        {
            Provider provider;
            Dictionary<string, object> result = new Dictionary<string, object>();
            try
            {
                provider = ProvidersList.Find(x => x.Id == idProv);
            }
            catch (Exception e)
            {
                throw e;
            }
            result.Add("Provider_Name", provider.Name);
            result.Add("Provider_Priority", provider.Priority);
            return result;
        }

        public int GetPriority(string id) {
            return ProvidersList.Where(x => x.Id == id).First().Priority;
        }

        public TimeInterval GetTimeInterval(string idProv,string idStock) {
            return ProvidersList.Where(x => x.Id == idProv).First().GetTimeInterval(idStock);
        }

        public List<string> GetId() {
            return ProvidersList.Select(x => x.Id).ToList();
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)ProvidersList).GetEnumerator();
        }

        public void Delete(string idProvider) {
            ProvidersList.RemoveAll(x => x.Id == idProvider);
        }

        public void Delete(string idProvider, string idStock) {
            var provider = ProvidersList.Find(x => x.Id == idProvider);
            provider.Delete(idStock);
        }
    }

    public class Provider: IEnumerable
    {
        public Stock this[int i] {
            get {
                return Stocks[i];
            }
        }

        public Stock this[string id] {
            get {
                return Stocks.Find(x => x.Id == id);
            }
        }

        private GenId GenId { set; get; }
        public string Id { get; private set; }
        public string Name { get; set; }
        public int Priority { get; set; }

        private List<Stock> Stocks { get; set; }

        public Provider(string id, string name, int priority)
        {
            Id = id;
            Name = name;
            Priority = priority;
            GenId = new GenId('A', -1, 1);
            Stocks = new List<Stock>();
        }

        public Provider(XmlNode x)
        {
            Stocks = new List<Stock>();
            Id = x.Attributes.GetNamedItem("id").Value;

            XmlNode xmlSet = x.SelectSingleNode("settings");
            GenId = new GenId(char.Parse(xmlSet.ChildNodes.Item(0).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(1).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(2).InnerText));

            Name = x.SelectSingleNode("name").InnerText;
            Priority = int.Parse(x.SelectSingleNode("priority").InnerText);

            foreach (XmlNode xNode in x.SelectSingleNode("stocks").ChildNodes)
            {
                Stocks.Add(new Stock(xNode));
            }

            Stocks.Sort((x1, x2) => x1.Name.CompareTo(x2.Name));
        }

        public string AddStock(string name, TimeInterval time)
        {
            string id = GenId.NexVal();
            Stocks.Add(new Stock(id, name, time));
            return id;
        }


        public string AddStock(string name, string time)
        {
            string id = GenId.NexVal();
            Stocks.Add(new Stock(id, name, time));
            return id;
        }

        public Stock GetStock(string id)
        {
            Stock stock = null;
            stock = Stocks.Find(x => x.Id == id);
            return stock;
        }

        public string[] GetStocksId()
        {
            return Stocks.Select(x => x.Id).ToArray();
        }

        public XmlNode GetXmlNode(XmlDocument xmlDocument)
        {
            XmlElement element = xmlDocument.CreateElement("provider");
            element.SetAttribute("id", Id);

            element.AppendChild(GenId.GetXmlNode(xmlDocument));

            XmlElement e = xmlDocument.CreateElement("name");
            e.InnerText = Name;
            element.AppendChild(e);
            e = xmlDocument.CreateElement("priority");
            e.InnerText = Priority.ToString();
            element.AppendChild(e);

            XmlElement stocksElement;
            element.AppendChild(stocksElement = xmlDocument.CreateElement("stocks"));
            foreach (var item in Stocks)
            {
                stocksElement.AppendChild(item.GetXmlNode(xmlDocument));
            }

            return element;
        }

        public Dictionary<string, object> GetValuesById(string idStock)
        {
            Stock stock;
            Dictionary<string, object> result = new Dictionary<string, object>();
            try
            {
                stock = Stocks.Find(x => x.Id == idStock);
            }
            catch (Exception e)
            {
                throw e;
            }
            result.Add("Stock_Name", stock.Name);
            result.Add("Stock_Time", stock.Time);
            return result;
        }

        public TimeInterval GetTimeInterval(string id) {
            return Stocks.Where(x => x.Id == id).First().Time;
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)Stocks).GetEnumerator();
        }

        public void Delete(string idStock) {
            Stocks.RemoveAll(x => x.Id == idStock);
        }
    }

    public class Stock
    {
        public string Name { get; set; }
        public string Id { get; private set; }
        public TimeInterval Time { get; set; }

        public Stock(string id, string name, TimeInterval time) {
            Name = name;
            Id = id;
            Time = time;
        }

        public Stock(string id, string name, string time)
        {
            Name = name;
            Id = id;
            Time = new TimeInterval(time);
        }

        public Stock(XmlNode xNode)
        {
            Id = xNode.Attributes.GetNamedItem("id").Value;
            Name = xNode.SelectSingleNode("name").InnerText;
            Time = new TimeInterval(xNode.SelectSingleNode("time").InnerText);
        }

        public XmlNode GetXmlNode(XmlDocument xmlDocument)
        {
            XmlElement element = xmlDocument.CreateElement("stock");
            element.SetAttribute("id", Id);

            XmlElement e = xmlDocument.CreateElement("name");
            e.InnerText = Name;
            element.AppendChild(e);

            e = xmlDocument.CreateElement("time");
            e.InnerText = Time.ToString();
            element.AppendChild(e);

            return element;
        }
    }
}