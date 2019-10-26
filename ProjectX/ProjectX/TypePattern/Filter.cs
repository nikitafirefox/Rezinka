using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ProjectX.Information;

namespace ProjectX.TypePattern
{
    public class Filter
    {
        // Товар
        public double MinCost = int.MinValue;
        public double MaxCost = int.MaxValue;

        public int MinCount = int.MinValue;
        public int MaxCount = int.MaxValue;
        //

        // Поставщик
        public List<string> ProvidersId = new List<string>();
        //

        //База знаний
        public List<string> BrandsId = new List<string>();
        public List<string> Seassons = new List<string>();

        public List<string> Widths = new List<string>();
        public List<string> Heights = new List<string>();
        public List<string> Diameters = new List<string>();
        public List<string> Spikes = new List<string>();
        public List<string> Acomadations = new List<string>();
        public List<string> Additions = new List<string>();

        //

        public List<Element> GetElements(List<Element> elements) {

            List<Element> res = new List<Element>();

            foreach (var item in elements)
            {


                if ((MinCost <= double.Parse(item.Price)) && (MaxCost >= double.Parse(item.Price)) 
                    && (MinCount <= int.Parse(item.Count)) && (MaxCount >= int.Parse(item.Count))
                    && (ProvidersId.Contains(item.ProvaiderId)) && (BrandsId.Contains(item.BrandId)) 
                    && (Seassons.Contains(item.Season)) && (Widths.Contains(item.Width)) 
                    && (Heights.Contains(item.Height)) && (Diameters.Contains(item.Diameter)) 
                    && (Spikes.Contains(item.Spikes)) && (!string.IsNullOrEmpty(item.Season))
                    && (Acomadations.Contains(item.Accomadation)) && (Additions.Contains(item.Addition))) {

                    res.Add(item);
                
                }

            }

            return res;
        }

    }
}
