using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;

namespace ProjectX.Dict
{

    public class Brand:IEnumerable
    {

        public string Id { get; private set; }
        public string Name { get; set; }
        public string Country { get; set; }
        public string Description { get; set; }

        public string RunFlatName { get; set; }

        private GenId IdGen { get; set; }
        private List<string> Variations { get; set; }
        private List<Model> Models { get; set; }
        private List<string> Images { get; set; }
        private List<string> Tags { get; set; }

        public Model this[string id] {
            get {
                return Models.Find(x => x.Id == id);
            }
        }

        public Brand(XmlNode x)
        {
            Images = new List<string>();
            Tags = new List<string>();
            Variations = new List<string>();
            Models = new List<Model>();

            Id = x.Attributes.GetNamedItem("id").Value;

            XmlNode xmlSet = x.SelectSingleNode("settings");
            IdGen = new GenId(char.Parse(xmlSet.ChildNodes.Item(0).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(1).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(2).InnerText));

            Name = x.SelectSingleNode("name").InnerText;
            Description = x.SelectSingleNode("description").InnerText;
            Country = x.SelectSingleNode("country").InnerText;
            RunFlatName = x.SelectSingleNode("runFlatName").InnerText;

            foreach (XmlNode xNode in x.SelectSingleNode("images").ChildNodes)
            {
                Images.Add(xNode.Attributes.GetNamedItem("src").Value);
            }

            foreach (XmlNode xNode in x.SelectSingleNode("variations").ChildNodes)
            {
                Variations.Add(xNode.InnerText);
            }

            XmlNode xmlTags = null;

            try {
                xmlTags = x.SelectSingleNode("tags");
            }
            catch{}

            if (xmlTags != null) {
                foreach (XmlNode xNode in xmlTags.ChildNodes) {
                    Tags.Add(xNode.InnerText);
                }
            }

            Tags.Sort();

            foreach (XmlNode xNode in x.SelectSingleNode("models").ChildNodes)
            {
                Models.Add(new Model(xNode));
            }

            Models.Sort((x1, x2) => x1.Name.CompareTo(x2.Name));
        }

        public Brand(string id, string name, string country, string description, string runFlatName)
        {
            Images = new List<string>();
            IdGen = new GenId('A', -1, 1);
            Variations = new List<string>();
            Models = new List<Model>();
            Tags = new List<string>();

            Id = id;
            Name = name;
            Country = country;
            Description = description;
            RunFlatName = runFlatName;
        }

        public override string ToString()
        {
            return "ID: " + Id + '\n'
                + "Name: " + Name + '\n'
                + "Country: " + Country + '\n'
                + "RunFlatName: " + RunFlatName + '\n'
                + "Description: " + Description + '\n'
                + "Image: \n\t" + String.Join("\n\t", Images) + '\n'
                + "Variations: \n\t" + String.Join("\n\t", Variations);
        }

        public XmlElement GetXmlNode(XmlDocument document)
        {
            XmlElement element = document.CreateElement("brand");
            element.SetAttribute("id", Id);

            element.AppendChild(IdGen.GetXmlNode(document));

            XmlElement e = document.CreateElement("name");
            e.InnerText = Name;
            element.AppendChild(e);
            e = document.CreateElement("country");
            e.InnerText = Country;
            element.AppendChild(e);
            e = document.CreateElement("runFlatName");
            e.InnerText = RunFlatName;
            element.AppendChild(e);

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

            XmlElement tagsElement;

            element.AppendChild(tagsElement = document.CreateElement("tags"));
            foreach (string item in Tags) {
                e = document.CreateElement("tag");
                e.InnerText = item;
                tagsElement.AppendChild(e);
            }

            XmlElement modelsElement;
            element.AppendChild(modelsElement = document.CreateElement("models"));
            foreach (Model item in Models)
            {
                modelsElement.AppendChild(item.GetXmlNode(document));
            }

            return element;
        }

        public string Add(string type, string name, string season, string description, bool commercial,
            string whileLetters, bool mudSnow)
        {
            string id = IdGen.NexVal();
            Models.Add(new Model(id, type, name, season, description, commercial, whileLetters, mudSnow));
            return id;
        }

        public string Add(string idModel, string width, string height, string diameter, string speedIndex,
            string loadIndex, string country, string tractionIndex, string temperatureIndex, string treadwearIndex,
            bool extraLoad, bool runFlat, string flangeProtection, string accomadation, bool spikes)
        {
            string id;
            try
            {
                id = Models.Find(x => x.Id == idModel).Add(width, height, diameter, speedIndex, loadIndex,
                country, tractionIndex, temperatureIndex, treadwearIndex, extraLoad, runFlat, flangeProtection, accomadation,spikes);
            }
            catch (ArgumentNullException)
            {
                throw new ArgumentNullException();
            }
            return id;
        }

        public IEnumerable<string> GetImages(string v2)
        {
            return Models.Find(x=> x.Id == v2).GetImages();
        }

        public IEnumerable<string> GetVariations() {
            return Variations;
        }

        public void AddStringValue(string value)
        {
            Variations.Add(value);
        }

        public void AddStringValue(string idModel, string value)
        {
            try
            {
                Models.Find(x => x.Id == idModel).AddStringValue(value);
            }
            catch (ArgumentNullException)
            {
                throw new ArgumentNullException();
            }
        }


        public List<Model> AnalysisModel(string parsingBufer,bool wordSearching, out List<string> variationStrings)
        {
            Models.Sort((x1, x2) => x1.Name.ToLower().CompareTo(x2.Name.ToLower()));
            List<Model> Resault = new List<Model>();
            variationStrings = new List<string>();
            string variation;
            foreach (var item in Models)
            {
                if (item.IsMatched(parsingBufer,wordSearching, out variation))
                {
                    bool isAdding = true;
                    List<string> resModels = new List<string>();

                    foreach (var var2 in variationStrings)
                    {
                        string minStr = variation.Length > var2.Length ? var2 : variation;
                        string maxStr = variation.Length >= var2.Length ? variation : var2;

                        if (Regex.IsMatch(maxStr, Regex.Escape(minStr), RegexOptions.IgnoreCase))
                        {
                            isAdding = false;
                            if (var2.Length == minStr.Length)
                            {
                                resModels.Add(var2);
                            }
                        }
                    }

                    foreach (var resModel in resModels)
                    {
                        Resault.RemoveAt(variationStrings.FindIndex(x => x == resModel));
                        variationStrings.Remove(resModel);
                    }

                    if (resModels.Count!=0 || isAdding)
                    {
                        Resault.Add(item);
                        variationStrings.Add(variation);
                    }
                }
            }
            return Resault;
        }

        public void Set(string name, string country, string description, string runFlatName)
        {
            Name = name;
            Country = country;
            Description = description;
            RunFlatName = runFlatName;
        }

        public void Set(string idModel, string name, string type, string season, string description, bool commercial,
            string whileLetters, bool mudSnow)
        {
            try
            {
                Models.Find(x => x.Id == idModel).Set(name, type, season, description, commercial, whileLetters,
                        mudSnow);
            }
            catch (ArgumentNullException)
            {
                throw new ArgumentNullException();
            }
        }

        public void Set(string idModel, string idMarking, string speedIndex,
            string loadIndex, string country, string tractionIndex, string temperatureIndex, string treadwearIndex,
            bool extraLoad, bool runFlat, string flangeProtection, string accomadation,bool spikes)
        {
            try
            {
                Models.Find(x => x.Id == idModel).Set(idMarking, speedIndex, loadIndex, country, tractionIndex, temperatureIndex,
                    treadwearIndex, extraLoad, runFlat, flangeProtection,accomadation,spikes);
            }
            catch (ArgumentNullException)
            {
                throw new ArgumentNullException();
            }
        }

        public void Delete(string idModel)
        {
            Models.Remove(Models.Find(x => x.Id == idModel));
        }

        public void Delete(string idModel, string idMarking)
        {
            Models.Find(x => x.Id == idModel).Delete(idMarking);
        }

        public void AddImage(string image)
        {
            Images.Add(image);
        }

        public void DeleteStringValue(string value)
        {
            Variations.Remove(value);
        }

        public void DeleteStringValue(string idModel, string value)
        {
            Models.Find(x => x.Id == idModel).DeleteStringValue(value);
        }



        public void AddImage(string idModel, string image)
        {
            try
            {
                Models.Find(x => x.Id == idModel).AddImage(image);
            }
            catch (ArgumentNullException)
            {
                throw new ArgumentNullException();
            }
        }

        public void DeleteImage(string image)
        {
            Images.Remove(image);
        }

        public void DeleteImage(string idModel, string image)
        {
            Models.Find(x => x.Id == idModel).DeleteImage(image);
        }

        public bool IsMatch(string buffer, bool wordSearching, out string variationString)
        {
            bool b = false;
            Variations.Sort((x, y) => y.Length.CompareTo(x.Length));
            variationString = "";
            foreach (var item in Variations)
            {
                string regeexStr = wordSearching ? "\\b" + Regex.Escape(item) + "\\b" : Regex.Escape(item);

                
                if (item.Contains(@"\") || item.Contains("+") || item.Contains("*") || item.Contains("?")
                        || item.Contains("|") || item.Contains(".") || item.Contains("$") || item.Contains("{") || item.Contains("[") ||
                        item.Contains("(") || item.Contains(")") || item.Contains("#")||item.Contains("@"))
                {
                    regeexStr = Regex.Escape(item);

                    if (wordSearching)
                    {

                        char first = item[0];
                        char last = item[item.Length - 1];


                        if (first != '\\' && first != '+' && first != '*' && first != '?'
                            && first != '|' && first != '.' && first != '$' && first != '{' && first != '[' &&
                            first != '(' && first != ')' && first != '#' && first != '@')
                        {
                            regeexStr = "\\b" + regeexStr;
                        }

                        if (last != '\\' && last != '+' && last != '*' && last != '?'
                            && last != '|' && last != '.' && last != '$' && last != '{' && last != '[' &&
                            last != '(' && last != ')' && last != '#' && last != '@')
                        {
                            regeexStr = regeexStr + "\\b";
                        }

                    }
                }

                

                if (Regex.IsMatch(buffer,regeexStr, RegexOptions.IgnoreCase))
                {
                    b = true;
                    variationString = item;
                    break;
                }
            }
            return b;
        }

        public Dictionary<string, object> GetValuesById(string idModel, string idMarking)
        {
            Dictionary<string, object> resault;
            try
            {
                Model model = Models.Find(x => x.Id == idModel);
                resault = model.GetValuesById(idMarking);
                resault.Add("Model_Name", model.Name);
                resault.Add("Model_Season", model.Season);
                resault.Add("Model_Description", model.Description);
                resault.Add("Model_Commercial", model.Commercial.ToString().Replace("True", "Да").Replace("False", "Нет"));
                resault.Add("Model_Type", model.Type);
                resault.Add("Model_MudSnow", model.MudSnow.ToString().Replace("True", "Да").Replace("False", "Нет"));
                resault.Add("Model_WhileLetters", model.WhileLetters);
            }
            catch (Exception e)
            {
                throw e;
            }
            return resault;
        }

        public void GetWidths(List<string> res)
        {
            foreach (Model item in Models)
            {
                item.GetWidths(res);
            }
            
        }

        public void GetHeights(List<string> res)
        {
            foreach (Model item in Models)
            {
                item.GetHeights(res);
            }
        }

        public void GetDiametrs(List<string> res)
        {
            foreach (Model item in Models)
            {
                item.GetDiametrs(res);
            }
        }

        public void GetSeassons(List<string> res)
        {
            foreach (Model item in Models)
            {
                res.Add(item.Season);
            }
        }

        public void GetAccomadation(List<string> res) {

            foreach (Model item in Models)
            {
                item.GetAccomadation(res);
            }
        }

        public void GetModelsName(List<string> res)
        {
            foreach (Model item in Models)
            {
                res.Add(item.Name);
            }
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)Models).GetEnumerator();
        }

        public void ChangeStringValue(string sOld,string sNew) {
            DeleteStringValue(sOld);
            AddStringValue(sNew);
        }

        public IEnumerable<string> GetImages() {
            return Images;
        }

        public IEnumerable<string> GetTags()
        {
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

        public void ChangeImage(string iOld, string iNew) {
            DeleteImage(iOld);
            AddImage(iNew);
        }
    }
}