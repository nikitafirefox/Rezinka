using System.Collections;
using System.Collections.Generic;

namespace ProjectX.ExcelParsing
{
    abstract public class Resault: IEnumerable
    {
        public string Message { get; set; }

        public int Information { get; set; }

        private List<string> Log { get; set; }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)Log).GetEnumerator();
        }

        public void AddLog(string str) {
            Log.Add(str);
        }

        public void AddLog(List<string> Lstr) {
            Log = Lstr;
        }

    }

    public class GResault : Resault
    {
        public string Id { get; private set; }

        public string Name { get; private set; }

        public string Addition { get; set; }

        public GResault(string id, string name,string addition, string mes, int inf)
        {
            Id = id;
            Name = name;
            Message = mes;
            Addition = addition;
            Information = inf;
        }
    }

    public class BResault : Resault
    {
        public BResault(string mes, int inf)
        {
            Message = mes;
            Information = inf;
        }
    }

    public class NResault : Resault
    {
        private Dictionary<string, string> KeyValuePairs { get; set; }

        public string BufferAfterParse;

        public NResault(string mes, Dictionary<string, string> keyValues, string bufferAfterParsing,int inf)
        {
            BufferAfterParse = bufferAfterParsing;
            KeyValuePairs = keyValues;
            Message = mes;
            Information = inf;
        }


    }
}