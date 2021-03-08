using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using N.Package.Excel.Internal;
using UnityEngine;

namespace N.Package.Excel.Model
{
    public class NExcelTable
    {
        public readonly List<NExcelColumn> Columns = new List<NExcelColumn>();
        public readonly List<NExcelRow> Rows = new List<NExcelRow>();

        public NExcelSimpleTable AsSimpleTable()
        {
            return NExcelUtils.AsSimpleTable(this);
        }

        public string AsJson()
        {
            var simpleRows = AsSimpleTable();
            return JsonUtility.ToJson(simpleRows, true);
        }
    }
}