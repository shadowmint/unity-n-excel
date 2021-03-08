using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using N.Package.Excel.Internal;
using UnityEngine;

namespace N.Package.Excel.Model
{
    public class NExcelSheet
    {
        public string Name { get; set; }
        public DataTable Table { get; set; }

        /// <summary>
        /// This is an opinionated table loader that loads an entire sheet as a table.
        /// The first non-empty row is loaded as the table headers, and then every row is added.
        /// If you don't like it, use the DataTable directly. Empty columns and rows are removed
        /// if 'dropBlank' is set.
        /// </summary>
        public NExcelTable LoadTable(bool dropBlank = true)
        {
            return NExcelUtils.LoadTable(this, dropBlank);
        }
    }
}