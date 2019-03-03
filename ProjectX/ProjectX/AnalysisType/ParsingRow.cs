using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectX.ExcelParsing
{
    public class ParsingRow:IEnumerable
    {
        public string ExcelRowIndex { get; set; }

        public string IdProvider { get; set; }

        public string ParsingBufer { get; private set;}

        public double Price { get; private set;}

        private List<ParsingCount> ParsingCounts { get; set; }

        public Resault Resault { get; set; }

        public ParsingRow(string parsingBufer,double price, string ERIndex, string idProvider) {
            ParsingCounts = new List<ParsingCount>();
            ParsingBufer = parsingBufer;
            Price = price;
            ExcelRowIndex = ERIndex;
            IdProvider = idProvider;
        }

        public override string ToString() {
            return ParsingBufer + '\t' + Price + "\n\t" +String.Join("\n\t",ParsingCounts);
        }

        public void AddCount(ParsingCount count) => ParsingCounts.Add(count);

        public IEnumerator GetEnumerator()
        {
            return ParsingCounts.GetEnumerator();
        }
    }

    public class ParsingCount
    {

        public string Id { set; get; }

        public int Count { get; set; }

        public ParsingCount(string id, int count) {
            Id = id;
            Count = count;
        }

        public override string ToString() => Id + '\t' + Count;
        
    }



}
