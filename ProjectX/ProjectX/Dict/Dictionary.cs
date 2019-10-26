using ProjectX.AnalysisType;
using ProjectX.ExcelParsing;
using ProjectX.Loger;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;

namespace ProjectX.Dict
{

    public class Dictionary:IEnumerable
    {

        public int AnalysisCount { get; set;}

        public int AnalysisMax { get; set; }

        public string AnalysisRowName { get; set; }

        public Brand this[string id] {
            get {
                return Brands.Find(x => x.Id == id);
            }
        }

        private List<Brand> Brands { get; set; }
        private GenId IdGen { get; set; }
        private readonly string pathXML;
        private SystemLoger Log { get; set; }
        private Thread AutoSaveThead { get; set; }
        private bool HaveChanged { get; set; }

        public Dictionary()
        {
            HaveChanged = false;
            Brands = new List<Brand>();
            pathXML = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\Dictionary.xml");
            Log = new SystemLoger(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\DictLog.txt"));

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
                Save("Создание документа");
            }
            else
            {
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

        private void InitDictionary()
        {
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

            Brands.Sort((x1,x2)=> x1.Name.CompareTo(x2.Name));
        }

        public void Save() => Save("Сохранение");

        private void Save(string massege)
        {
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

        private void WriteLog(string message)
        {
            new Thread(new ParameterizedThreadStart(Log.WriteLog))
            {
                Priority = ThreadPriority.Lowest
            }.Start(message);
        }

        public string Add(string name, string country, string description, string runFlatName)
        {
            Stopwatch sw = new Stopwatch();
            string id;
            lock (this)
            {
                sw.Start();
                id = IdGen.NexVal();
                Brands.Add(new Brand(id, name, country, description, runFlatName));
                HaveChanged = true;
                sw.Stop();
            }
            WriteLog("Добавлен брэнд " + id + "(" + sw.ElapsedMilliseconds + "мс)");
            return id;
        }

        public string Add(string idBrand, string type, string name, string season, string description, bool commercial,
            string whileLetters, bool mudSnow)
        {
            Stopwatch sw = new Stopwatch();
            string id;
            lock (this)
            {
                sw.Start();
                try
                {
                    id = Brands.Find(x => x.Id == idBrand).Add(type, name, season, description,
                        commercial, whileLetters, mudSnow);
                }
                catch (ArgumentNullException)
                {
                    throw new Exception("Ошибка словаря(3)");
                }
                HaveChanged = true;
                sw.Stop();
            }
            WriteLog("Добавлена  модель " + idBrand + "-" + id + "(" + sw.ElapsedMilliseconds + "мс)");
            return id;
        }

        public IEnumerable<string> GetImages(string v1, string v2)
        {
            return Brands.Find(x => x.Id == v1).GetImages(v2);
        }

        public string Add(string idBrand, string idModel, string width, string height, string diameter,
            string speedIndex, string loadIndex, string country, string tractionIndex, string temperatureIndex,
            string treadwearIndex, bool extraLoad, bool runFlat, string flangeProtection, string accomadation,bool spikes)
        {
            Stopwatch sw = new Stopwatch();
            string id;
            lock (this)
            {
                sw.Start();
                try
                {
                    id = Brands.Find(x => x.Id == idBrand).Add(idModel, width, height, diameter, speedIndex,
                        loadIndex, country, tractionIndex, temperatureIndex, treadwearIndex, extraLoad,
                        runFlat, flangeProtection,accomadation,spikes);
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

        public void AddStringValue(string idBrand, string value)
        {
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
            WriteLog("Добавлена  вариация к " + idBrand + "(" + sw.ElapsedMilliseconds + "мс)");
        }

        public void AddStringValue(string idBrand, string idModel, string value)
        {
            Stopwatch sw = new Stopwatch();
            lock (this)
            {
                sw.Start();
                try
                {
                    Brands.Find(x => x.Id == idBrand).AddStringValue(idModel, value);
                }
                catch (ArgumentNullException)
                {
                    throw new Exception("Ошибка словаря(3)");
                }
                HaveChanged = true;
                sw.Stop();
            }
            WriteLog("Добавлена  вариация к " + idBrand + "-" + idModel + "(" + sw.ElapsedMilliseconds + "мс)");
        }


        public void AddImage(string idBrand, string image)
        {
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
                    Brands.Find(x => x.Id == idBrand).Set(name, country, description, runFlatName);
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
            string whileLetters, bool mudSnow)
        {
            Stopwatch sw = new Stopwatch();
            lock (this)
            {
                sw.Start();
                try
                {
                    Brands.Find(x => x.Id == idBrand).Set(idModel, name, type, season, description, commercial, whileLetters,
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
            bool extraLoad, bool runFlat, string flangeProtection, string accomadation,bool spikes)
        {
            Stopwatch sw = new Stopwatch();
            lock (this)
            {
                sw.Start();
                try
                {
                    Brands.Find(x => x.Id == idBrand).Set(idModel, idMarking, speedIndex, loadIndex, country, tractionIndex, temperatureIndex,
                    treadwearIndex, extraLoad, runFlat, flangeProtection,accomadation, spikes);
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

        public void Delete(string idBrand)
        {
            lock (this)
            {
                Brands.Remove(Brands.Find(x => x.Id == idBrand));
            }
        }

        public void Delete(string idBrand, string idModel)
        {
            lock (this)
            {
                Brands.Find(x => x.Id == idBrand).Delete(idModel);
            }
        }

        public void Delete(string idBrand, string idModel, string idMarking)
        {
            lock (this)
            {
                Brands.Find(x => x.Id == idBrand).Delete(idModel, idMarking);
            }
        }

        public void DeleteStringValue(string idBrand, string value)
        {
            lock (this)
            {
                Brands.Find(x => x.Id == idBrand).DeleteStringValue(value);
            }
        }

        public void DeleteStringValue(string idBrand, string idModel, string value)
        {
            lock (this)
            {
                Brands.Find(x => x.Id == idBrand).DeleteStringValue(idModel, value);
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
            catch (Exception e)
            {
                throw e;
            }
            return resault;
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

        private void AutoSave()
        {
            while (true)
            {
                try
                {
                    Thread.Sleep((int)6E+5);
                    if (HaveChanged)
                    {
                        Save("Автосохранение");
                    }
                }
                catch (ThreadInterruptedException)
                {
                    break;
                }
            }
        }

        public void Close() => AutoSaveThead.Interrupt();

        public List<Brand> AnalysisBrand(string buffer, bool wordSearching, out List<string> variationsStrings)
        {
            List<Brand> Resault = new List<Brand>();
            variationsStrings = new List<string>();
            string variation = "";
            foreach (var item in Brands)
            {
                if (item.IsMatch(buffer, wordSearching, out variation))
                {
                    bool isAdding = true;
                    List<string> resBrands = new List<string>();
                    foreach (var var2 in variationsStrings)
                    {
                        string minStr = variation.Length > var2.Length ? var2 : variation;
                        string maxStr = variation.Length >= var2.Length ? variation : var2;

                        if (Regex.IsMatch(maxStr, Regex.Escape(minStr), RegexOptions.IgnoreCase))
                        {
                            isAdding = false;
                            if (var2.Length == minStr.Length)
                            {
                                resBrands.Add(var2);
                            }
                        }
                    }
                    foreach (var resBrand in resBrands)
                    {
                        Resault.RemoveAt(variationsStrings.FindIndex(x => x == resBrand));
                        variationsStrings.Remove(resBrand);
                    }
                    if (resBrands.Count!=0 || isAdding)
                    {
                        Resault.Add(item);
                        variationsStrings.Add(variation);
                    }
                }
            }
            return Resault;
        }

        public List<ParsingRow> Analysis(ref List<ParsingRow> parsingRows)
        {
            return Analysis(ref parsingRows, false,false);
        }

        public List<ParsingRow> Analysis(ref List<ParsingRow> parsingRows, bool deepSearch)
        {
            return Analysis(ref parsingRows, deepSearch, false);
        }

        public List<ParsingRow> Analysis(ref List<ParsingRow> parsingRows, bool deepSearch, bool wordSearching) {
            return Analysis(ref parsingRows, deepSearch, wordSearching, false);
        }

       

        public List<ParsingRow> Analysis(ref List<ParsingRow> parsingRows, bool deepSearch,bool wordSearching,bool autoComplate)
        {
            List<ParsingRow> Resault = new List<ParsingRow>();

            ProviderRegulars providerRegulars = new ProviderRegulars();

            AnalysisMax = parsingRows.Count;

            AnalysisCount = 0;

            AnalysisRowName = "";

            foreach (var item in parsingRows)
            {

                if (item.ExcelRowIndex == "102"|| item.ExcelRowIndex == "336") {
                    int itrtr = 0;
                }


                AnalysisCount++;

                AnalysisRowName = item.ParsingBufer;

                List<string> lStr = new List<string>();
                int countLog = 0;
                
                string parsBuf = item.ParsingBufer;

                lStr.Add(++countLog + ") Определяем наличие маркировки ");

                int countMarking = providerRegulars.CountMarking(item.IdProvider, item.ParsingBufer);

                lStr.Add(++countLog + ") Маркировок найдено: " + countMarking);
                
                if (countMarking == 1)
                {


                    List<string> variationsStringsBrands = new List<string>();
                    List<string> variationsStringsModels = new List<string>();
                    string sov_variation;

                    lStr.Add(++countLog + ") Попытка парсинга производителя ");

                    List<Brand> brands = AnalysisBrand(parsBuf,wordSearching, out variationsStringsBrands);

                    string logStr = ++countLog + ") Производителей найдено: " + brands.Count;

                    for (int i = 0; i < brands.Count; i++) {
                        logStr += "\nID: " + brands[i].Id + "  Название: " + brands[i].Name + " Определен как: " + variationsStringsBrands[i] + '\n';
                    }

                    lStr.Add(logStr);

                    string name = "Товар";
                    string id = "";

                    bool already_find_model = false;

                    Brand findedBrand = null;
                    Model findedModel = null;
                    if (brands.Count == 0 && deepSearch)
                    {
                        lStr.Add(++countLog + ") Функиция углубленного поиска включена ");
                        brands = Brands;
                        lStr.Add(++countLog + ") Производителей взято: " + brands.Count);
                    }
                    else if (brands.Count == 0 && !deepSearch)
                    {
                        lStr.Add(++countLog + ") Функиция углубленного поиска  отключена");

                        Dictionary<string, string> keyValuePairs = providerRegulars.GetDictionary(item.IdProvider, parsBuf, out string str);
                        string mes = "Не найден производитель";

                        lStr.Add(++countLog + ") Требуются уточнения");
                        item.Resault = new NResault(mes, keyValuePairs, str.Trim(),4);
                        item.Resault.AddLog(lStr);
                        continue;

                    }

                    foreach (Brand brand in brands)
                    {
                        lStr.Add(++countLog + ") Попытка парсинга модели у производителя " + brand.Name);

                        List<string> variationsStrings = new List<string>();
                        List<Model> models = brand.AnalysisModel(parsBuf,wordSearching, out variationsStrings);

                        logStr = ++countLog + ") Найдено моделей: "+ models.Count;
                        for (int i = 0; i < models.Count; i++)
                        {
                            logStr += "\nID: " + brand.Id + "-" + models[i].Id + "  Название: " + models[i].Name + " Определен как: " + variationsStrings[i];
                        }
                        lStr.Add(logStr);

                        if (models.Count == 1 && !(already_find_model))
                        {
                              

                            findedModel = models.First();
                            findedBrand = brand;
                            already_find_model = true;
                            variationsStringsModels = variationsStrings;

                            
                        }
                        else if (models.Count != 0)
                        {
                            lStr.Add(++countLog + ") Вторая попытка парсинга модели");

                            lStr.Add(++countLog + ") Попытка парсинга дополнительных значений");
                            Dictionary<string,string> dictionary =providerRegulars.GetDictionary(item.IdProvider, parsBuf, out string newBuf);
                            
                            logStr = ++countLog + ") Найдено значений: " + dictionary.Count;

                            foreach (var val in dictionary) {
                                logStr += "\n" + val.Key + " : " + val.Value;
                            }
                            
                            lStr.Add(logStr);

                            lStr.Add(++countLog + ") Строка после парсинга значений: " + newBuf);


                            lStr.Add(++countLog + ") Попытка парсинга модели у производителя " + brand.Name);
                            

                            models = brand.AnalysisModel(newBuf,wordSearching, out variationsStrings);
                           

                            
                            logStr = ++countLog + ") Найдено моделей: " + models.Count;
                            for (int i = 0; i < models.Count; i++)
                            {
                                logStr += "\nID: " + brand.Id + "-" + models[i].Id + "  Название: " + models[i].Name + " Определен как: " + variationsStrings[i];
                            }
                            lStr.Add(logStr);
                            


                            if (models.Count == 1 && !(already_find_model))
                            {
                                already_find_model = true;
                                findedModel = models.First();
                                findedBrand = brand;
                                already_find_model = true;
                                variationsStringsModels = variationsStrings;
                                continue;
                            }



                            already_find_model = false;

                            lStr.Add(++countLog + ") Требуются уточнения");
                            item.Resault = new NResault("Найдено " + models.Count +  " моделей", dictionary, newBuf.Trim(), 5);
                            item.Resault.AddLog(lStr);
                            break;
                        }
                    }

                    if (!already_find_model)
                    {
                        if (item.Resault == null)
                        {
                            


                            string mes = "Не найдена модель";

                            lStr.Add(++countLog + ") Попытка парсинга дополнительных значений");

                            Dictionary<string, string> keyValuePairs = providerRegulars.GetDictionary(item.IdProvider, parsBuf, out string str);

                            logStr = ++countLog + ") Найдено значений: " + keyValuePairs.Count;

                            foreach (var val in keyValuePairs)
                            {
                                logStr += "\n" + val.Key + " : " + val.Value;
                            }

                            lStr.Add(logStr);

                            lStr.Add(++countLog + ") Строка после парсинга значений: " + str);

                            int inf = 3;
                            if (brands.Count == 0 || brands.Count == Brands.Count) {
                                mes = "Не найден производитель";
                                inf = 4;
                                lStr.Add(++countLog + ") Производитель не найден ");
                            }
                            else {
                                lStr.Add(++countLog + ") Модель не найдена ");

                                variationsStringsBrands.Sort((x, y) => y.Length.CompareTo(x.Length));

                                logStr = ++countLog + ") Удаление всех вхождений производителя";

                                foreach (string variation in variationsStringsBrands)
                                {
                                    logStr += "\nИсходная строка: " + str;
                                    logStr += " Найденная вариация: " + variation;
                                    sov_variation = Regex.Match(parsBuf, Regex.Escape(variation), RegexOptions.IgnoreCase).Value;
                                    logStr += " Найденное вхождение: " + sov_variation;
                                    if (sov_variation != "")
                                    {
                                        str = str.Replace(sov_variation, "");
                                        logStr += "(успешно) результат: " + str;
                                        
                                    }
                                    logStr += "(ошибка удаления)";
                                }

                                lStr.Add(logStr);
                            }



                            

                          

                            lStr.Add(++countLog + ") Требуются уточнения");
                            item.Resault = new NResault(mes, keyValuePairs, str.Trim(), inf) {FindBrandId = brands.Count == 1 ? brands.First().Id : null };
                            item.Resault.AddLog(lStr);
                        }
                        continue;
                    }

                    logStr = ++countLog + ") Удаление всех вхождений модели";

                    variationsStringsModels.Sort((x, y) => y.Length.CompareTo(x.Length));

                    foreach (string variation in variationsStringsModels)
                    {
                        logStr += "\nИсходная строка: " + parsBuf;
                        logStr += " Найденная вариация: " + variation;
                        sov_variation = Regex.Match(parsBuf, Regex.Escape(variation), RegexOptions.IgnoreCase).Value;
                        logStr += " Найденное вхождение: " + sov_variation;
                        if (sov_variation != "")
                        {
                            parsBuf = parsBuf.Replace(sov_variation, "");
                            logStr += "(успешно) результат: " + parsBuf;
                        }
                        logStr += "(ошибка удаления)";
                    }

                    lStr.Add(logStr);


                    logStr = ++countLog + ") Удаление всех вхождений производителя";

                    variationsStringsBrands.Sort((x, y) => y.Length.CompareTo(x.Length));

                    foreach (string variation in variationsStringsBrands)
                    {
                        logStr += "\nИсходная строка: " + parsBuf;
                        logStr += " Найденная вариация: " + variation;
                        sov_variation = Regex.Match(parsBuf, Regex.Escape(variation), RegexOptions.IgnoreCase).Value;
                        logStr += " Найденное вхождение: " + sov_variation;
                        if (sov_variation != "")
                        {
                            parsBuf = parsBuf.Replace(sov_variation, "");
                            logStr += "(успешно) результат: " + parsBuf;
                        }
                        logStr += "(ошибка удаления)";
                    }

                    lStr.Add(logStr);



                    id = findedBrand.Id;
                    name += " " + findedBrand.Name;
                    id += "-" + findedModel.Id;
                    name += " " + findedModel.Name;

                    lStr.Add(++countLog + ") Попытка парсинга дополнительных значений");

                    Dictionary<string, string> keyValues = providerRegulars.GetDictionary(item.IdProvider, parsBuf, out string s);

                    logStr = ++countLog + ") Найдено значений: " + keyValues.Count;

                    foreach (var val in keyValues)
                    {
                        logStr += "\n" + val.Key + " : " + val.Value;
                    }

                    lStr.Add(logStr);


                    lStr.Add(++countLog + ") Строка после парсинга значений: " + s);

                    if (s.Trim() != "")
                    {
                       
                        lStr.Add(++countLog + ") Требуются уточнения");
                        item.Resault = new NResault("При парсинги остались символы в исходной строке", keyValues, s.Trim(), 2) {
                            FindBrandId = findedBrand.Id,
                            FindModelId = findedModel.Id
                        };
                        item.Resault.AddLog(lStr);
                        continue;
                    }



                    string width;
                    string height;
                    string diameter;
                    bool commercial = false;
                    string indexSpeed = "";
                    string loadIndex = "";
                    bool extraLoad = false;
                    string season = "";
                    bool runFlat = false;
                    string countryBrand = "";
                    string countryMarking = "";
                    string tractionIndex = "";
                    string temperatureIndex = "";
                    string treadwearIndex = "";
                    string flangeProtection = "";
                    string whileLetters = "";
                    bool mudSnow = false;
                    string runFlatName = "";
                    string type = "";
                    string accomadation = "";
                    string additional = "";
                    bool spikes = false;

                    try
                    {
                        width = keyValues["width"];
                        height = keyValues["height"];
                        diameter = keyValues["diameter"];
                        indexSpeed = keyValues["speedIndex"];
                        loadIndex = keyValues["loadIndex"];
                    }
                    catch
                    {
                        lStr.Add(++countLog + ") ошибка, не хватает значениия(ий): ширины, высоты, диаметра, " +
                            "индекса скороти, индекса нагрузки");


                        lStr.Add(++countLog + ") Завершено с ошибкой");
                        item.Resault = new BResault(item.ExcelRowIndex + "Ошибка при полученнии маркировки", 6);
                        item.Resault.AddLog(lStr);

                        continue;
                    }

                    if (keyValues.ContainsKey("commercial"))
                    {
                        commercial = true;
                        if (autoComplate && !findedModel.Commercial)
                        {
                            findedModel.Commercial = commercial;
                        }
                    }

                    if (keyValues.ContainsKey("extraLoad"))
                    {
                        extraLoad = true;
                        
                    }

                    if (keyValues.ContainsKey("season"))
                    {
                        season = keyValues["season"];

                        if ( string.IsNullOrEmpty(findedModel.Season)  && autoComplate)
                        {
                            findedModel.Season = season;
                        }
                    }

                    if (keyValues.ContainsKey("runFlat"))
                    {
                        runFlat = true;
                    }

                    if (keyValues.ContainsKey("countryBrand"))
                    {
                        countryBrand = keyValues["countryBrand"];

                        if (string.IsNullOrEmpty(findedBrand.Country) && autoComplate)
                        {
                            findedBrand.Country = countryBrand;
                        }
                    }

                    if (keyValues.ContainsKey("countryMarking"))
                    {
                        countryMarking = keyValues["countryMarking"];

                    }

                    if (keyValues.ContainsKey("tractionIndex"))
                    {
                        tractionIndex = keyValues["tractionIndex"];
                    }

                    if (keyValues.ContainsKey("temperatureIndex"))
                    {
                        temperatureIndex = keyValues["temperatureIndex"];
                    }

                    if (keyValues.ContainsKey("treadwearIndex"))
                    {
                        treadwearIndex = keyValues["treadwearIndex"];
                    }

                    if (keyValues.ContainsKey("whileLetters"))
                    {
                        whileLetters = keyValues["whileLetters"];

                        if (string.IsNullOrEmpty(findedModel.WhileLetters) && autoComplate)
                        {
                            findedModel.WhileLetters = whileLetters;
                        }
                    }

                    if (keyValues.ContainsKey("mudSnow"))
                    {
                        mudSnow = true;

                        if (!findedModel.MudSnow && autoComplate)
                        {
                            findedModel.MudSnow = mudSnow;
                        }
                        
                    }

                    if (keyValues.ContainsKey("runFlatName"))
                    {
                        runFlatName = keyValues["runFlatName"];

                        if (string.IsNullOrEmpty(findedBrand.RunFlatName) && autoComplate)
                        {
                            findedBrand.RunFlatName = runFlatName;
                        }
                    }

                    if (keyValues.ContainsKey("type"))
                    {
                        type = keyValues["type"];

                        if (string.IsNullOrEmpty(findedModel.Type) && autoComplate)
                        {
                            findedModel.Type = type;
                        }

                        
                    }

                    if (keyValues.ContainsKey("accomadation"))
                    {
                        accomadation = keyValues["accomadation"];
                    }

                    if (keyValues.ContainsKey("additional"))
                    {
                        additional = keyValues["additional"];
                    }

                    if (keyValues.ContainsKey("spikes"))
                    {
                        spikes = true;
                    }

                    if (keyValues.ContainsKey("flangeProtection"))
                    {
                        flangeProtection = keyValues["flangeProtection"];
                    }


                    lStr.Add(++countLog + ") Попытка парсинга маркировки");

                    Marking marking = findedModel.SearchMarking(width, height, diameter,loadIndex, indexSpeed, accomadation,spikes,runFlat, out bool isContainMarking);
                    if (isContainMarking)
                    {
                        id += "-" + marking.Id;
                        name += " " + marking.Width + "/" + marking.Height + "R" + marking.Diameter;
                        if (marking.Accomadation != "") { name += " акомадация" + marking.Accomadation; }

                        if (string.IsNullOrEmpty(marking.Country) && autoComplate) {
                            marking.Country = countryMarking;
                        }

                        if (string.IsNullOrEmpty(marking.TractionIndex) && autoComplate)
                        {
                            marking.TractionIndex = tractionIndex;
                        }

                        if (string.IsNullOrEmpty(marking.TemperatureIndex) && autoComplate)
                        {
                            marking.TemperatureIndex = temperatureIndex;
                        }

                        if (string.IsNullOrEmpty(marking.TreadwearIndex) && autoComplate)
                        {
                            marking.TreadwearIndex = treadwearIndex;
                        }

                        if (string.IsNullOrEmpty(marking.FlangeProtection) && autoComplate)
                        {
                            marking.FlangeProtection = flangeProtection;
                        }

                        if (!marking.ExtraLoad && autoComplate)
                        {
                            marking.ExtraLoad = extraLoad;
                        }

                        lStr.Add(++countLog + ") Маркировка найдена\n Определена как: " + id + "\nЗаписана в базу как: " + name);

                        lStr.Add(++countLog + ") Завершено успещно");

                        item.Resault = new GResault(id, name,additional, "Существующий",1);
                        item.Resault.AddLog(lStr);

                    }
                    else
                    {
                        id += "-" + findedModel.Add(width, height, diameter, indexSpeed, loadIndex, 
                            countryMarking, tractionIndex, temperatureIndex, treadwearIndex, extraLoad, 
                            runFlat, flangeProtection, accomadation,spikes);
                        name += " " + width + "/" + height + "R" + diameter;
                        if (accomadation != "") { name += " акомадация " + accomadation; }


                        lStr.Add(++countLog + ") Маркировка добавлена\n Определена как: " + id + "\nЗаписана в базу как: " + name);

                        lStr.Add(++countLog + ") Завершено успещно");

                        item.Resault = new GResault(id, name,additional, "Новый", 0);

                        item.Resault.AddLog(lStr);
                    }
                }
                else if (countMarking == 0)
                {
                    lStr.Add(++countLog + ") Завершено с ошибкой");

                    item.Resault = new BResault(" Отсутствие маркировки ", 7);
                    item.Resault.AddLog(lStr);
                }
                else
                {
                    lStr.Add(++countLog + ") Завершено с ошибкой");

                    item.Resault = new BResault(" Найдено " + countMarking + " маркировок ", 8);
                    item.Resault.AddLog(lStr);
                }
            }

            return parsingRows;
        }

        public List<string> GetBrandsId() {
            return Brands.Select(x => x.Id).ToList();
        }

        public List<string> GetWidths() {
            List<string> res = new List<string>();
            foreach (Brand item in Brands)
            {
                item.GetWidths(res);
            }
            res.Sort();
            return res.Distinct().ToList();
        }

        public List<string> GetHeights()
        {
            List<string> res = new List<string>();
            foreach (Brand item in Brands)
            {
                item.GetHeights(res);
            }
            res.Sort();
            return res.Distinct().ToList();
        }

        public List<string> GetDiametrs()
        {
            List<string> res = new List<string>();
            foreach (Brand item in Brands)
            {
                item.GetDiametrs(res);
            }
            res.Sort();
            return res.Distinct().ToList();
        }

        public List<string> GetSeassons()
        {
            List<string> res = new List<string>();
            foreach (Brand item in Brands)
            {
                item.GetSeassons(res);
            }
            res.Sort();
            return res.Distinct().ToList();
        }

        public List<string> GetAccomadation()
        {
            List<string> res = new List<string>();
            foreach (Brand item in Brands)
            {
                item.GetAccomadation(res);
            }
            res.Sort();
            return res.Distinct().ToList();
        }

        public List<string> GetModelsName() {
            List<string> res = new List<string>();
            foreach (Brand item in Brands)
            {
                item.GetModelsName(res);
            }
            res.Sort();
            return res.Distinct().ToList();
        }

        //********************************************************//

        private struct M
        {
            public string name;
            public string img;
            public List<string> var;

            public void Init()
            {
                var = new List<string>();
            }

            public void Change(string name, string img)
            {
                this.name = name;
                this.img = img;
            }
        }

        private struct P
        {
            public string name;
            public string var;
            public List<M> ms;

            public void Init()
            {
                ms = new List<M>();
            }
        }

        public void GetItemFromTxt(string path)
        {
            List<P> ps = new List<P>();

            using (StreamReader sr = new StreamReader(path))
            {
                string buf;
                P CurentP = new P();
                CurentP.Init();

                while ((buf = sr.ReadLine()) != null)
                {
                    string[] vs = buf.Split('\t');

                    if (vs[0] == "")
                    {
                        M CurentM;
                        if (CurentP.ms.FindIndex(x => x.name == vs[2]) != -1)
                        {
                            CurentM = CurentP.ms.Find(x => x.name == vs[2]);
                        }
                        else
                        {
                            CurentM = new M();
                            CurentM.Init();
                            CurentM.img = vs[4];
                            CurentM.name = vs[2];
                            CurentP.ms.Add(CurentM);
                        }
                        CurentM.var.Add(vs[1]);
                    }
                    else
                    {
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

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)Brands).GetEnumerator();
        }

        //*******************************************************************//
    }
}