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
        public int Priority { get; set; }

        public string IdProduct { get; set; }
        public string IdRow { get; set; }

        public string Price { get; set; }
        public string PriceForOne { get; set; }
        public string PriceForTwo { get; set; }
        public string Count { get; set; }
        public string Addition { get; set; }

        public string ProvaiderId { get; set; }
        public string TimeTransit { get; set; }

        public string BrandId { get; set; }
        public string BrandName { get; set; }
        public string BrandCountry { get; set; }
        public string BrandDescription { get; set; }
        public string RunFlatName { get; set; }

        public string ModelName { get; set; }
        public string Season { get; set; }
        public string Type { get; set; }
        public string ModelDescription { get; set; }
        public string Commercial { get; set; }
        public string WhileLetters { get; set; }
        public string MudSnow { get; set; }

        public string Width { get; set; }
        public string Height { get; set; }
        public string Diameter { get; set; }
        public string SpeedIndex { get; set; }
        public string LoadIndex { get; set; }
        public string MarkingCountry { get; set; }
        public string TractionIndex { get; set; }
        public string TemperatureIndex { get; set; }
        public string TreadwearIndex { get; set; }
        public string ExtraLoad { get; set; }
        public string RunFlat { get; set; }
        public string FlangeProtection { get; set; }
        public string Accomadation { get; set; }
        public string Spikes { get; set; }

        public string Date { get; set; }
        public string Time { get; set; }

        public int TimeFromTo { get; set; }

        public bool FreeFitting { get; set; }

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


                Brand brand = dictionary[item.IdProduct.Split('-')[0]];
                Model model = brand[item.IdProduct.Split('-')[1]];
                Marking marking = model[item.IdProduct.Split('-')[2]];


                Provider provider = providers[item.IdPosition.Split('-')[0]];
                Stock stock = provider[item.IdPosition.Split('-')[1]];

                int timeFromTo = stock.Time.Days + stock.Time.Month * 30 + stock.Time.Weeks * 7;

                Element element = new Element()
                {
                    Addition = item.Addition,
                    Count = item.Count.ToString(),
                    Price = item.TotalPrice.ToString(),
                    PriceForOne = item.TotalPriceForOne.ToString(),
                    PriceForTwo = item.TotalPriceForTwo.ToString(),
                    BrandId = item.IdProduct.Split('-')[0],
                    ProvaiderId = item.IdPosition.Split('-')[0],
                    TimeTransit = stock.Time.ToString(),
                    Priority = provider.Priority,
                    BrandName = brand.Name,
                    BrandCountry = brand.Country,
                    BrandDescription = brand.Description,
                    RunFlatName = brand.RunFlatName,

                    ModelName = model.Name,
                    Season = model.Season,
                    ModelDescription = model.Description,
                    Commercial = model.Commercial.ToString().Replace("True", "Да").Replace("False", "Нет"),
                    Type = model.Type,
                    MudSnow = model.MudSnow.ToString().Replace("True", "Да").Replace("False", "Нет"),
                    WhileLetters = model.WhileLetters,

                    Width = marking.Width,
                    Height = marking.Height,
                    Accomadation = marking.Accomadation,
                    Diameter = marking.Diameter,
                    ExtraLoad = marking.ExtraLoad.ToString().Replace("True", "Да").Replace("False", "Нет"),
                    FlangeProtection = marking.FlangeProtection,
                    LoadIndex = marking.LoadIndex,
                    MarkingCountry = marking.Country,
                    RunFlat = marking.RunFlat.ToString().Replace("True", "Да").Replace("False", "Нет"),
                    SpeedIndex = marking.SpeedIndex,
                    Spikes = marking.Spikes.ToString().Replace("True", "Да").Replace("False", "Нет"),
                    TemperatureIndex = marking.TemperatureIndex,
                    TractionIndex = marking.TractionIndex,
                    TreadwearIndex = marking.TreadwearIndex,

                    Date = date,
                    Time = time,

                    IdProduct = item.IdProduct,
                    IdRow = item.IdRow,

                    TimeFromTo = timeFromTo,

                    FreeFitting = item.Count >= 4 && ((brand.FreeFitting && int.Parse(marking.Diameter.Trim()) >= brand.FittingDiameter)
                    || (model.FreeFitting && int.Parse(marking.Diameter.Trim()) >= model.FittingDiameter)),

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
                if (double.Parse(item.Price) <= 0) {
                    continue;
                }

                var el = res.Find(x => x.IdProduct == item.IdProduct && item.Addition == x.Addition);

                if (el == null)
                {
                    res.Add(item);
                }
                else {
                    if (el.ProvaiderId == item.ProvaiderId) {

                        if (double.Parse(el.Price) <= double.Parse(item.Price))
                        {
                            el.Price = item.Price;
                            el.PriceForTwo = item.PriceForTwo;
                            el.PriceForOne = item.PriceForOne;

                            el.Count = (int.Parse(el.Count) + int.Parse(item.Count)).ToString();
                            
                        }

                        if (new TimeInterval(el.TimeTransit) > new TimeInterval(item.TimeTransit)) {
                            el.TimeTransit = item.TimeTransit;
                            el.TimeFromTo = item.TimeFromTo;
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


                case "Priority":
                    return Priority.ToString();


                case "PriceCount":
                    if (int.Parse(Count) == 1) {
                        return "one";
                    }
                    if (int.Parse(Count) < 4) {
                        return "two";
                    }
                    return "four";




                default: return "System error";
            }
        }
    }
}
