using System;
using System.Collections.Generic;

namespace N.Package.Excel.Model
{
    [System.Serializable]
    public class NExcelRow
    {
        public bool isBlank;
        public Dictionary<string, string> schema;
        public Dictionary<string, object> values;
    }
}