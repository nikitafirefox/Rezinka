using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ProjectX.TireFitting
{
    public class FittingDict: IEnumerable
    {
        public double RunFlatCost { get; set; }
        private List<Fitting> fittings;
        private List<AddingService> addings;

        public Fitting this[int diameter] {
            get {
                Fitting fitting = fittings.Find(x => x.Diameter == diameter);
                return fitting;
            }
        }
    
        private readonly string pathXML;

        public FittingDict() {

            fittings = new List<Fitting>();
            pathXML = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\Fitting.xml");

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

                RunFlatCost = 0;

                Save();

            }
            else {
                InitFittingsDict();
            }

        }

        public void InitFittingsDict() {

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(pathXML);
            XmlElement xroot = xmlDocument.DocumentElement;

            RunFlatCost = double.Parse(xroot.SelectSingleNode("runFlatCost").InnerText);

            XmlNode xmlNode = xroot.GetElementsByTagName("fittings").Item(0);
            foreach (XmlNode x in xmlNode.ChildNodes)
            {
                fittings.Add(new Fitting(x));
            }

            fittings.Sort((x1, x2) => x2.Diameter.CompareTo(x2.Diameter));

            xmlNode = xroot.GetElementsByTagName("addings").Item(0);
            foreach (XmlNode x in xmlNode.ChildNodes)
            {
                addings.Add(new AddingService(x));
            }

        }

        public void Save()
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(pathXML);
            XmlElement xroot = xmlDocument.DocumentElement;
            xroot.RemoveAll();

            XmlElement e = xmlDocument.CreateElement("runFlatCost");
            e.InnerText = RunFlatCost.ToString();
            xroot.AppendChild(e);

            XmlElement fittingsElement;
            xroot.AppendChild(fittingsElement = xmlDocument.CreateElement("fittings"));
            foreach (var item in fittings)
            {
                fittingsElement.AppendChild(item.GetXmlNode(xmlDocument));
            }

            XmlElement addingsElement;
            xroot.AppendChild(addingsElement = xmlDocument.CreateElement("addings"));
            foreach (var item in addings)
            {
                addingsElement.AppendChild(item.GetXmlNode(xmlDocument));
            }

            xmlDocument.Save(pathXML);
        }

        public IEnumerable<AddingService> GetAddings() {
            return addings;
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)fittings).GetEnumerator();
        }

        public void AddFitting(int diameter, double autoCost, double outRiderCost) {
            fittings.Add(new Fitting(diameter, autoCost, outRiderCost));
        }

        public void AddAddings(string name, double cost) {
            addings.Add(new AddingService(name, cost));
        }

        public void DeleteFitting(int diameter) {
            fittings.Remove(fittings.Find(x => x.Diameter == diameter));
        }

        public void DeleteAdding(string name) {
            addings.Remove(addings.Find(x => x.ServiceName == name));
        }

        public void ChangeFitting(int diameter, double autoCost, double outRiderCost) {
            Fitting fitting = fittings.Find(x => x.Diameter == diameter);
            if (fitting != null) {
                fitting.AutoCost = autoCost;
                fitting.OutRiderCost = outRiderCost;
            }
        }

        public void ChangeAdding(string name, double cost) {
            AddingService adding = addings.Find(x => x.ServiceName == name);
            if (adding != null) {
                adding.Cost = cost;
            }
        }

    }

    public class Fitting {

        public int Diameter { get; set; }
        public double AutoCost { get; set; }
        public double OutRiderCost { get; set; }

        public Fitting(int diameter, double autoCost, double outRiderCost) {
            Diameter = diameter;
            AutoCost = autoCost;
            OutRiderCost = outRiderCost;
        }

        public Fitting(XmlNode xNode){
            Diameter = int.Parse(xNode.SelectSingleNode("diameter").InnerText);
            AutoCost = double.Parse(xNode.SelectSingleNode("autoCost").InnerText);
            OutRiderCost = double.Parse(xNode.SelectSingleNode("outRiderCost").InnerText);
        }

        public XmlNode GetXmlNode(XmlDocument xmlDocument) {
            XmlElement element = xmlDocument.CreateElement("fitting");

            XmlElement e = xmlDocument.CreateElement("diameter");
            e.InnerText = Diameter.ToString();
            element.AppendChild(e);

            e = xmlDocument.CreateElement("autoCost");
            e.InnerText = AutoCost.ToString();
            element.AppendChild(e);

            e = xmlDocument.CreateElement("outRiderCost");
            e.InnerText = OutRiderCost.ToString();
            element.AppendChild(e);

            return element;
        }

    }

    public class AddingService {

        public string ServiceName { get; set; }
        public double Cost { get; set; }

        public AddingService(string serviceName, double cost) {
            ServiceName = serviceName;
            Cost = cost;
        }

        public AddingService(XmlNode xNode){
            ServiceName = xNode.SelectSingleNode("serviceName").InnerText;
            Cost = double.Parse(xNode.SelectSingleNode("cost").InnerText);
        }

        public XmlNode GetXmlNode(XmlDocument xmlDocument) {

            XmlElement element = xmlDocument.CreateElement("adding");

            XmlElement e = xmlDocument.CreateElement("serviceName");
            e.InnerText = ServiceName.ToString();
            element.AppendChild(e);

            e = xmlDocument.CreateElement("cost");
            e.InnerText = Cost.ToString();
            element.AppendChild(e);

            return element;

        }
        

    }
}
