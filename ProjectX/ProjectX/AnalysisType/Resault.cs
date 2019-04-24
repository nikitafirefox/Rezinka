using System.Collections.Generic;

namespace ProjectX.ExcelParsing
{
    abstract public class Resault
    {
        public string Message { get; set; }
    }

    public class GResault : Resault
    {
        public string Id { get; private set; }

        public string Name { get; private set; }

        public GResault(string id, string name, string mes)
        {
            Id = id;
            Name = name;
            Message = mes;
        }
    }

    public class BResault : Resault
    {
        public BResault(string mes)
        {
            Message = mes;
        }
    }

    public class NResault : Resault
    {
        private Dictionary<string, string> KeyValuePairs { get; set; }

        public string BufferAfterParse;

        public NResault(string mes, Dictionary<string, string> keyValues, string bufferAfterParsing)
        {
            BufferAfterParse = bufferAfterParsing;
            KeyValuePairs = keyValues;
            Message = mes;
        }
    }
}