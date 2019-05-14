using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ProjectX.Information;
using ProjectX.Dict;
using ProjectX.DataBase;

namespace ProjectX.TypePattern
{
    public class Element
    {
        public int Priority;

        public string IdProduct;
        public string IdRow;

        public string Price;
        public string Count;
        public string Addition;

        public string ProvaiderId;
        public string TimeTransit;

        public string BrandId;
        public string BrandName;
        public string BrandCountry;
        public string BrandDescription;
        public string RunFlatName;

        public string ModelName;
        public string Season;
        public string Type;
        public string ModelDescription;
        public string Commercial;
        public string WhileLetters;
        public string MudSnow;

        public string Width;
        public string Height;
        public string Diameter;
        public string SpeedIndex ;
        public string LoadIndex;
        public string MarkingCountry;
        public string TractionIndex ;
        public string TemperatureIndex ;
        public string TreadwearIndex ;
        public string ExtraLoad ;
        public string RunFlat ;
        public string FlangeProtection ;
        public string Accomadation;
        public string Spikes;

        public string Date;
        public string Time;

        public List<string> Images = new List<string>();

        public static List<Element> GetElements() {
            List<Element> res = new List<Element>();

            string date = DateTime.Now.ToShortDateString();
            string time = DateTime.Now.ToShortTimeString();

            DBase dBase = new DBase();
            Providers providers = new Providers();
            Dictionary dictionary = new Dictionary();

            foreach (DBRow item in dBase)
            {
                Dictionary<string, object> valuePairsProv = providers.GetValuesById(
                    item.IdPosition.Split('-')[0], 
                    item.IdPosition.Split('-')[1]
                    );
                Dictionary<string, object> valuePairsDict = dictionary.GetValuesById(
                    item.IdProduct.Split('-')[0],
                    item.IdProduct.Split('-')[1],
                    item.IdProduct.Split('-')[2]
                    );

                Element element = new Element()
                {
                    Addition = item.Addition,
                    Count = item.Count.ToString(),
                    Price = item.TotalPrice.ToString(),
                    BrandId = item.IdProduct.Split('-')[0],
                    ProvaiderId = item.IdPosition.Split('-')[0],
                    TimeTransit = valuePairsProv["Stock_Time"].ToString(),
                    Priority = (int)valuePairsProv["Provider_Priority"],
                    BrandName = valuePairsDict["Brand_Name"].ToString(),
                    BrandCountry = valuePairsDict["Brand_Country"].ToString(),
                    BrandDescription = valuePairsDict["Brand_Description"].ToString(),
                    RunFlatName = valuePairsDict["Brand_RunFlatName"].ToString(),
                    ModelName = valuePairsDict["Model_Name"].ToString(),
                    Season = valuePairsDict["Model_Season"].ToString(),
                    ModelDescription = valuePairsDict["Model_Description"].ToString(),
                    Commercial = valuePairsDict["Model_Commercial"].ToString(),
                    Type = valuePairsDict["Model_Type"].ToString(),
                    MudSnow = valuePairsDict["Model_MudSnow"].ToString(),
                    WhileLetters = valuePairsDict["Model_WhileLetters"].ToString(),

                    Width = valuePairsDict["Marking_Width"].ToString(),
                    Height = valuePairsDict["Marking_Height"].ToString(),
                    Accomadation = valuePairsDict["Marking_Accomadation"].ToString(),
                    Diameter = valuePairsDict["Marking_Diameter"].ToString(),
                    ExtraLoad = valuePairsDict["Marking_ExtraLoad"].ToString(),
                    FlangeProtection = valuePairsDict["Marking_FlangeProtection"].ToString(),
                    LoadIndex = valuePairsDict["Marking_LoadIndex"].ToString(),
                    MarkingCountry = valuePairsDict["Marking_Country"].ToString(),
                    RunFlat = valuePairsDict["Marking_RunFlat"].ToString(),
                    SpeedIndex = valuePairsDict["Marking_SpeedIndex"].ToString(),
                    Spikes = valuePairsDict["Marking_Spikes"].ToString(),
                    TemperatureIndex = valuePairsDict["Marking_TemperatureIndex"].ToString(),
                    TractionIndex = valuePairsDict["Marking_TractionIndex"].ToString(),
                    TreadwearIndex = valuePairsDict["Marking_TreadwearIndex"].ToString(),

                    Date = date,
                    Time = time,

                    IdProduct = item.IdProduct,
                    IdRow = item.IdRow
                
                    

                };

                element.Images = new List<string>();

                foreach (string str in dictionary.GetImages(item.IdProduct.Split('-')[0],
                    item.IdProduct.Split('-')[1])) {
                    element.Images.Add(str);
                }

                res.Add(element);
            }

            dictionary.Close();
            return res;
        }


        public static List<Element> Distinct(List<Element> elements) {
            List<Element> res = new List<Element>();

            elements.Sort((x, y) => x.Priority.CompareTo(y.Priority));
            foreach (var item in elements)
            {

                var el = res.Find(x => x.IdProduct == item.IdProduct && item.Addition == x.Addition);

                if (el == null)
                {
                    res.Add(item);
                }
                else {
                    if (el.ProvaiderId == item.ProvaiderId) {
                        if (double.Parse(el.Price) < double.Parse(item.Price))
                        {
                            el.Price = item.Price;
                            el.Count = (int.Parse(item.Count) + int.Parse(el.Count)).ToString();
                        }

                        if (new TimeInterval(el.TimeTransit) > new TimeInterval(item.TimeTransit)) {
                            el.TimeTransit = item.TimeTransit;
                        }
                        
                    }
                }
            }

            return res;
        }

        public string GetParam(string param)
        {
            switch (param) {

                case "Accomadation":
                case "Ac":
                    return Accomadation;

                case "Addition":
                case "Add":
                    return Addition;

                case "BrandCountry":
                case "BCountry":
                    return BrandCountry;

                case "BrandDescription":
                case "BDesc":
                    return BrandDescription;

                case "BrandName":
                case "BName":
                    return BrandName;

                case "Commercial":
                case "C":
                    return Commercial;

                case "Count":
                case "Ост":
                    return Count;

                case "Date":
                    return Date;

                case "Diameter":
                case "D":
                case "d":
                    return Diameter;

                case "ExtraLoad":
                case "XL":
                case "xl":
                    return ExtraLoad;

                case "FlangeProtection":
                case "FP":
                    return FlangeProtection;

                case "Height":
                case "H":
                case "h":
                    return Height;

                case "LoadIndex":
                case "LI":
                    return LoadIndex;

                case "MarkingCountry":
                case "MCountry":
                    return MarkingCountry;

                case "ModelDescription":
                case "MDesc":
                    return ModelDescription;

                case "ModelName":
                case "MName":
                    return ModelName;

                case "MudSnow":
                case "M+S":
                    return MudSnow;

                case "Price":
                case "$":
                    return Price;

                case "RunFlat":
                case "RF":
                    return RunFlat;

                case "RunFlatName":
                    return RunFlatName;

                case "Season":
                    return Season;

                case "SpeedIndex":
                case "SI":
                    return SpeedIndex;

                case "Spikes":
                case "Шип":
                    return Spikes;

                case "TemperatureIndex":
                case "Temper":
                    return TemperatureIndex;

                case "Time":
                    return Time;

                case "TimeTransit":
                case "Дост":
                    return TimeTransit;

                case "TractionIndex":
                case "Tract":
                    return TractionIndex;

                case "TreadwearIndex":
                case "Treadwear":
                    return TreadwearIndex;

                case "Type":
                    return Type;

                case "WhileLetters":
                case "WL":
                    return WhileLetters;

                case "Width":
                case "W":
                case "w":
                    return Width;

                case "idProvider":
                case "id_Поставщика":
                    return ProvaiderId;
                









                default: return "System error";
            }
        }
    }
}
