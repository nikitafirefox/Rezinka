using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using ProjectX.Loger;
using System.Text;
using System.Diagnostics;
using System.Threading;
using ProjectX.ExcelParsing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ProjectX.Dict
{
    public class Dictionary
    {
        private List<Brand> Brands { get; set;}
        private GenId IdGen { get; set; }
        private readonly string pathXML;
        private SystemLoger Log { get; set;}
        private Thread AutoSaveThead { get; set;}
        private bool HaveChanged { get; set;}

        public Dictionary()
        {
            HaveChanged = false;
            Brands = new List<Brand>();
            pathXML = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\Dictionary.xml");
            Log = new SystemLoger(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\DictLog.txt"));

            if (!File.Exists(pathXML)) {

                string dataDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data");
                if (!Directory.Exists(dataDirectory)) { Directory.CreateDirectory(dataDirectory); }
                IdGen = new GenId('A', -1, 1);
                FileStream fs = new FileStream(pathXML,FileMode.Create);
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
                Save("Создание документа");



            }
            else {



                Stopwatch sw = new Stopwatch();
                sw.Start();

                InitDictionary();

                sw.Stop();
                WriteLog("Инициализация (" + sw.ElapsedMilliseconds + "мс)");
            }

            AutoSaveThead = new Thread(new ThreadStart(AutoSave));
            AutoSaveThead.Priority = ThreadPriority.BelowNormal;
            AutoSaveThead.Start();

        }

        private void InitDictionary() {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(pathXML);
            XmlElement xroot = xmlDocument.DocumentElement;

            XmlNode xmlSet = xroot.SelectSingleNode("settings");
            IdGen = new GenId(Char.Parse(xmlSet.ChildNodes.Item(0).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(1).InnerText),
                int.Parse(xmlSet.ChildNodes.Item(2).InnerText));

            XmlNode xmlNode = xroot.GetElementsByTagName("brands").Item(0);
            foreach (XmlNode x in xmlNode.ChildNodes)
            {
                Brands.Add(new Brand(x));
            }
        }

        public void Save() => Save("Сохранение");

        private void Save(string massege) {

            
            Stopwatch sw = new Stopwatch();
            

            lock (this)
            {
                sw.Start();

                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(pathXML);
                XmlElement xroot = xmlDocument.DocumentElement;
                xroot.RemoveAll();
                xroot.AppendChild(IdGen.GetXmlNode(xmlDocument));
                XmlElement brandsElement;
                xroot.AppendChild(brandsElement = xmlDocument.CreateElement("brands"));
                foreach (Brand item in Brands)
                {
                    brandsElement.AppendChild(item.GetXmlNode(xmlDocument));
                }
                xmlDocument.Save(pathXML);
                HaveChanged = false;

                sw.Stop();
            }
            
            WriteLog(massege + "(" + sw.ElapsedMilliseconds + "мс)");
        }

        private void WriteLog(string message) {
            new Thread(new ParameterizedThreadStart(Log.WriteLog))
            {
                Priority = ThreadPriority.Lowest
            }.Start(message);
        }

        public string Add(string name, string country, string description, string runFlatName) {
            Stopwatch sw = new Stopwatch();
            string id;
            lock (this) {
                sw.Start();
                id = IdGen.NexVal();
                Brands.Add(new Brand(id, name, country, description, runFlatName));
                HaveChanged = true;
                sw.Stop();
            }
            WriteLog("Добавлен брэнд "+ id + "(" + sw.ElapsedMilliseconds + "мс)");
            return id;
        }

        public string Add(string idBrand, string type, string name, string season, string description, bool commercial,
            string whileLetters, bool mudSnow) {
            Stopwatch sw = new Stopwatch();
            string id;
            lock (this) {
                sw.Start();
                try
                {
                    id = Brands.Find(x => x.Id == idBrand).Add(type,name,season,description,
                        commercial,whileLetters,mudSnow);
                }
                catch (ArgumentNullException) {
                    throw new Exception("Ошибка словаря(3)");
                }
                HaveChanged = true;
                sw.Stop();
            }
            WriteLog("Добавлена  модель " +idBrand +"-"+ id + "(" + sw.ElapsedMilliseconds + "мс)");
            return id;
        }

        public string Add(string idBrand,string idModel, string width, string height, string diameter, 
            string speedIndex, string loadIndex, string country, string tractionIndex, string temperatureIndex,
            string treadwearIndex, bool extraLoad, bool runFlat, string flangeProtection) {
            Stopwatch sw = new Stopwatch();
            string id;
            lock (this)
            {
                sw.Start();
                try
                {
                    id = Brands.Find(x => x.Id == idBrand).Add(idModel, width, height, diameter, speedIndex, 
                        loadIndex, country, tractionIndex, temperatureIndex, treadwearIndex, extraLoad, 
                        runFlat, flangeProtection);
                }
                catch (ArgumentNullException)
                {
                    throw new Exception("Ошибка словаря(3)");
                }
                HaveChanged = true;
                sw.Stop();
            }
            WriteLog("Добавлена  маркеровка " + idBrand + "-" + idModel + "-" 
                + id + "(" + sw.ElapsedMilliseconds + "мс)");
            return id;
        }

        public void AddStringValue(string idBrand, string value) {
            Stopwatch sw = new Stopwatch();
            lock (this)
            {
                sw.Start();
                try
                {
                    Brands.Find(x => x.Id == idBrand).AddStringValue(value);
                }
                catch (ArgumentNullException)
                {
                    throw new Exception("Ошибка словаря(3)");
                }
                HaveChanged = true;
                sw.Stop();
            }
            WriteLog("Добавлена  вариация к " + idBrand  + "(" + sw.ElapsedMilliseconds + "мс)");
        }

        public void AddStringValue(string idBrand,string idModel, string value) {
            Stopwatch sw = new Stopwatch();
            lock (this)
            {
                sw.Start();
                try
                {
                    Brands.Find(x => x.Id == idBrand).AddStringValue(idModel,value);
                }
                catch (ArgumentNullException)
                {
                    throw new Exception("Ошибка словаря(3)");
                }
                HaveChanged = true;
                sw.Stop();
            }
            WriteLog("Добавлена  вариация к " + idBrand +"-" + idModel + "(" + sw.ElapsedMilliseconds + "мс)");
        }

        public void AddStringValue(string idBrand, string idModel, string idMarking, string value)
        {
            Stopwatch sw = new Stopwatch();
            lock (this)
            {
                sw.Start();
                try
                {
                    Brands.Find(x => x.Id == idBrand).AddStringValue(idModel, idMarking, value);
                }
                catch (ArgumentNullException)
                {
                    throw new Exception("Ошибка словаря(3)");
                }
                HaveChanged = true;
                sw.Stop();
            }
            WriteLog("Добавлена  аккомадация к " + idBrand + "-" + idModel +"-" + idMarking + 
                "(" + sw.ElapsedMilliseconds + "мс)");
        }

        public void AddImage(string idBrand, string image) {
            Stopwatch sw = new Stopwatch();
            lock (this)
            {
                sw.Start();
                try
                {
                    Brands.Find(x => x.Id == idBrand).AddImage(image);
                }
                catch (ArgumentNullException)
                {
                    throw new Exception("Ошибка словаря(3)");
                }
                HaveChanged = true;
                sw.Stop();
            }
            WriteLog("Добавлено изображение к " + idBrand + "(" + sw.ElapsedMilliseconds + "мс)");
        }

        public void AddImage(string idBrand, string idModel, string image)
        {
            Stopwatch sw = new Stopwatch();
            lock (this)
            {
                sw.Start();
                try
                {
                    Brands.Find(x => x.Id == idBrand).AddImage(idModel, image);
                }
                catch (ArgumentNullException)
                {
                    throw new Exception("Ошибка словаря(3)");
                }
                HaveChanged = true;
                sw.Stop();
            }
            WriteLog("Добавлено изображение к " + idBrand + "-" + idModel + "(" + sw.ElapsedMilliseconds + "мс)");
        }

        public void Set(string idBrand, string name, string country, string description,
            string runFlatName)
        {
            Stopwatch sw = new Stopwatch();
            lock (this)
            {
                sw.Start();
                try
                {
                    Brands.Find(x => x.Id == idBrand).Set(name,country,description,runFlatName);
                }
                catch (ArgumentNullException)
                {
                    throw new Exception("Ошибка словаря(3)");
                }
                HaveChanged = true;
                sw.Stop();
            }
            WriteLog("Изменен бренд " + idBrand + "(" + sw.ElapsedMilliseconds + "мс)");
        }

        public void Set(string idBrand, string idModel, string name, string type, string season, string description, bool commercial,
            string whileLetters, bool mudSnow) {
            Stopwatch sw = new Stopwatch();
            lock (this)
            {
                sw.Start();
                try
                {
                    Brands.Find(x => x.Id == idBrand).Set(idModel, name,type,season,description,commercial,whileLetters,
                        mudSnow);
                }
                catch (ArgumentNullException)
                {
                    throw new Exception("Ошибка словаря(3)");
                }
                HaveChanged = true;
                sw.Stop();
            }
            WriteLog("Изменена модель " + idBrand + "-" + idModel + "(" + sw.ElapsedMilliseconds + "мс)");
        }


        public void Set(string idBrand, string idModel, string idMarking, string speedIndex,
            string loadIndex, string country, string tractionIndex, string temperatureIndex, string treadwearIndex,
            bool extraLoad, bool runFlat, string flangeProtection)
        {
            Stopwatch sw = new Stopwatch();
            lock (this)
            {
                sw.Start();
                try
                {
                    Brands.Find(x => x.Id == idBrand).Set(idModel,idMarking, speedIndex, loadIndex, country, tractionIndex, temperatureIndex,
                    treadwearIndex, extraLoad, runFlat, flangeProtection);
                }
                catch (ArgumentNullException)
                {
                    throw new Exception("Ошибка словаря(3)");
                }
                HaveChanged = true;
                sw.Stop();
            }
            WriteLog("Изменена маркировка " + idBrand + "-" + idModel + "-" + idMarking + "(" + sw.ElapsedMilliseconds + "мс)");
        }

        public void Delete(string idBrand) {
            lock (this) {
                Brands.Remove(Brands.Find(x=> x.Id == idBrand));
            }

        }

        public void Delete(string idBrand, string idModel) {
            lock (this) {
                Brands.Find(x => x.Id == idBrand).Delete(idModel);
            }
        }

        public void Delete(string idBrand, string idModel, string idMarking) {
            lock (this) {
                Brands.Find(x => x.Id == idBrand).Delete(idModel,idMarking);
            }
        }

        public void DeleteStringValue(string idBrand, string value) {
            lock (this) {
                Brands.Find(x => x.Id == idBrand).DeleteStringValue(value);
            }

        }

        public void DeleteStringValue(string idBrand, string idModel, string value)
        {
            lock (this)
            {
                Brands.Find(x => x.Id == idBrand).DeleteStringValue(idModel,value);
            }

        }

        public Dictionary<string, object> GetValuesById(string idBrand, string idModel, string idMarking)
        {
            Dictionary<string, object> resault;
            try
            {
                Brand brand = Brands.Find(x => x.Id == idBrand);
                resault = brand.GetValuesById(idModel, idMarking);
                resault.Add("Brand_Name", brand.Name);
                resault.Add("Brand_Country", brand.Country);
                resault.Add("Brand_RunFlatName", brand.RunFlatName);
                resault.Add("Brand_Description", brand.Description);
            }
            catch (Exception e) {
                throw e;
            }
            return resault;
        }

        public void DeleteStringValue(string idBrand, string idModel, string idMarking, string value)
        {
            lock (this)
            {
                Brands.Find(x => x.Id == idBrand).DeleteStringValue(idModel, idMarking, value);
            }

        }

        public void DeleteImage(string idBrand, string image)
        {
            lock (this)
            {
                Brands.Find(x => x.Id == idBrand).DeleteImage(image);
            }

        }

        public void DeleteImage(string idBrand, string idModel, string image)
        {
            lock (this)
            {
                Brands.Find(x => x.Id == idBrand).DeleteImage(idModel, image);
            }

        }

        private void AutoSave() {
            while (true) {
                try
                {
                    Thread.Sleep((int)6E+5);
                    if (HaveChanged)
                    {
                        Save("Автосохранение");
                    }
                }catch (ThreadInterruptedException){
                    break;
                }
            }
        }

        public void Close() => AutoSaveThead.Interrupt();

        public List<Brand> AnalysisBrand(string buffer) {
            List<Brand> Resault = new List<Brand>();
            foreach (var item in Brands)
            {
                if (item.IsMatch(buffer)) {
                    Resault.Add(item);
                }
            }
            return Resault;
        }

        public  List<ParsingRow> Analysis(ref List<ParsingRow> parsingRows)
        {
            

            List<ParsingRow> Resault = new List<ParsingRow>();

            foreach (var item in parsingRows)
            {
                if (Regex.IsMatch(item.ParsingBufer, @"[0-9]{3}/[0-9]{2}R[0-9]{2}", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(item.ParsingBufer, @"[0-9]{3}/[0-9]{2}ZR[0-9]{2}", RegexOptions.IgnoreCase) ||
                     Regex.IsMatch(item.ParsingBufer, @"[0-9]{3}/[0-9]{2}/[0-9]{2}", RegexOptions.IgnoreCase))
                {
                    List<Brand> brands = AnalysisBrand(item.ParsingBufer);
                    string name = "Товар";
                    if (brands.Count == 1)
                    {
                        string id = brands.First().Id;
                        name += " " + brands.First().Name;
                        List<Model> models = brands.First().AnalysisModel(item.ParsingBufer);
                        if (models.Count == 1)
                        {
                            id += "-" + models.First().Id;
                            name += " " + models.First().Name;
                            List <Marking> markings = models.First().AnalysisMarking(item.ParsingBufer);
                            if (markings.Count == 0)
                            {
                                
                                string mark = null;
                                string width = null;
                                string height = null;
                                string diameter = null;
                                if (Regex.IsMatch(item.ParsingBufer, @"[0-9]{3}/[0-9]{2}R[0-9]{2}", RegexOptions.IgnoreCase))
                                {
                                    mark = Regex.Match(item.ParsingBufer, @"[0-9]{3}/[0-9]{2}R[0-9]{2}", RegexOptions.IgnoreCase).Value;
                                    width = Regex.Match(mark, @"[0-9]{3}", RegexOptions.IgnoreCase).Value;
                                    height = Regex.Match(mark, @"/[0-9]{2}", RegexOptions.IgnoreCase).Value.Replace("/", "");
                                    diameter = Regex.Match(mark, @"R[0-9]{2}", RegexOptions.IgnoreCase).Value.ToUpper().Replace("R", "");
                                }
                                else if (Regex.IsMatch(item.ParsingBufer, @"[0-9]{3}/[0-9]{2}ZR[0-9]{2}", RegexOptions.IgnoreCase))
                                {
                                    mark = Regex.Match(item.ParsingBufer, @"[0-9]{3}/[0-9]{2}ZR[0-9]{2}", RegexOptions.IgnoreCase).Value;
                                    width = Regex.Match(mark, @"[0-9]{3}", RegexOptions.IgnoreCase).Value;
                                    height = Regex.Match(mark, @"/[0-9]{2}", RegexOptions.IgnoreCase).Value.Replace("/", "");
                                    diameter = Regex.Match(mark, @"R[0-9]{2}", RegexOptions.IgnoreCase).Value.ToUpper().Replace("R", "");
                                }
                                else
                                {
                                    mark = Regex.Match(item.ParsingBufer, @"[0-9]{3}/[0-9]{2}/[0-9]{2}", RegexOptions.IgnoreCase).Value;
                                    width = Regex.Match(mark, @"[0-9]{3}", RegexOptions.IgnoreCase).Value;
                                    height = Regex.Matches(mark, @"/[0-9]{2}", RegexOptions.IgnoreCase)[0].Value.Replace("/", "");
                                    diameter = Regex.Matches(mark, @"/[0-9]{2}", RegexOptions.IgnoreCase)[1].Value.ToUpper().Replace("/", "");
                                }

                                id += "-" + models.First().Add(width, height, diameter, "", "", "", "", "", "", false, false, "");
                                name += " " + width + "/" + height + "R" + diameter;
                                item.Resault = new GResault(id, name, "Новый");
                            }
                            else if (markings.Count == 1)
                            {
                                id += "-" + markings.First().Id;
                                name += " " + markings.First().Width + "/" + markings.First().Height + "R" + markings.First().Diameter;
                                item.Resault = new GResault(id, name, "Существующий");
                            }
                            else
                            {
                                item.Resault = new BResault(item.ExcelRowIndex + ") Найдена несколько маркировок (error 3)"); 
                            }

                        }

                    }
                }
                else
                {
                    item.Resault = new BResault(item.ExcelRowIndex + ") Отсутствие маркировки (error 0)");
                }
            }

            return parsingRows;
        }

       
        
        
        
        
        //********************************************************//

        struct M {
           public string name;
           public string img;
           public List<string> var;

           public void Init() {
                var = new List<string>();
            }

            public void Change(string name,string img) {
                this.name = name;
                this.img = img;
            }
        }

        struct P
        {
           public string name;
           public string var;
           public List<M> ms;

           public void Init() {
                ms = new List<M>();
            }
        }

        internal void GetItemFromTxt(string path) {
            List<P> ps = new List<P>();

            using (StreamReader sr = new StreamReader(path)) {

                string buf;
                P CurentP = new P();
                CurentP.Init();

                while ((buf = sr.ReadLine()) != null) {

                    string[] vs = buf.Split('\t');

                    if (vs[0] == "")
                    {
                        M CurentM;
                        if (CurentP.ms.FindIndex(x => x.name == vs[2]) != -1)
                        {
                            CurentM = CurentP.ms.Find(x => x.name == vs[2]);
                        }
                        else {
                            CurentM = new M();
                            CurentM.Init();
                            CurentM.img = vs[4];
                            CurentM.name = vs[2];
                            CurentP.ms.Add(CurentM);
                        }
                        CurentM.var.Add(vs[1]);

                    }
                    else {
                        CurentP = new P();
                        CurentP.Init();
                        CurentP.var = vs[0];
                        CurentP.name = vs[1];
                        ps.Add(CurentP);
                    }

                }
            }


            foreach (var B in ps)
             {
                 string idB = Add(B.name, "", "", "");
                 AddStringValue(idB, B.var);
                 foreach (M Mod in B.ms)
                 {
                     string idM = Add(idB, "", Mod.name, "", "", false, "", false);
                     AddImage(idB, idM, Mod.img);
                     foreach (string str in Mod.var)
                     {
                         AddStringValue(idB, idM, str);
                     }
                 }
             }
        }
        //*******************************************************************//

    }
}
