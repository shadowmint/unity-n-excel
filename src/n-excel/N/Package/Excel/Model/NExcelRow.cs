using System;
using System.Collections.Generic;

namespace N.Package.Excel.Model
{
    public class NExcelRow
    {
        public bool IsBlank { get; set; }
        public Dictionary<string, object> Values { get; set; }
    }
}