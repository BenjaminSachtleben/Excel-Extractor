using System;
using System.Collections.Generic;

namespace ExcelExtractor
{
    public class EEmodel
    {
        public string source = string.Empty;
        public string target = string.Empty;
        public string cells = string.Empty;
        public List<string> sourcefiles = new List<string>();
        public List<string> cellNames = new List<string>();
        public Dictionary<string,List<string>> cellValues = new Dictionary<string,  List<string>>();
    }
}
