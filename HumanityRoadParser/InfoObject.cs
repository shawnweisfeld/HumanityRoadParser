using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HumanityRoadParser
{
    public class InfoObject
    {
        public InfoObject()
        {
            ParsedValues = new List<string>();
        }

        public int Row { get; set; }

        public int Column { get; set; }

        public string State { get; set; }

        public string AgencyType { get; set; }

        public string Name { get; set; }

        public string ColumnName { get; set; }

        public string Value { get; set; }

        public ValueType ValueType { get; set; }

        public List<string> ParsedValues { get; set; }
    }
}
