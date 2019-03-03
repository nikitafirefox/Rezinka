using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectX.ExcelParsing
{
    abstract public class Resault
    {
        public string Message { get; set;}


    }

    public class GResault : Resault
    {
        public string Id { get; private set; }

        public string Name { get; private set; }

        public GResault(string id,string name, string mes)
        {
            Id = id;
            Name = name;
            Message = mes;
        }

    }

    public class BResault : Resault
    {
        public BResault(string mes) {
            Message = mes;
        }
    }

    public class NResault : Resault
    {
        
    }

}
