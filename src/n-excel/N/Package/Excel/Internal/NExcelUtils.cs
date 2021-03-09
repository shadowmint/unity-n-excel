using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using N.Package.Excel.Model;

namespace N.Package.Excel.Internal
{
    public class NExcelUtils
    {
        public static NExcelSheet[] LoadSheetsFrom(DataSet result)
        {
            var sheets = new List<NExcelSheet>();
            for (var i = 0; i < result.Tables.Count; i++)
            {
                var table = result.Tables[i];
                sheets.Add(new NExcelSheet()
                {
                    Name = table.TableName,
                    Table = table,
                });
            }

            return sheets.ToArray();
        }

        /// <summary>
        /// This is an opinionated table loader that loads an entire sheet as a table.
        /// The first non-empty row is loaded as the table headers, and then every row is added.
        /// If you don't like it, use the DataTable directly. Empty columns and rows are removed
        /// if 'dropBlank' is set.
        /// </summary>
        public static NExcelTable LoadTable(NExcelSheet sheet, bool dropBlank = true)
        {
            var table = new NExcelTable();
            var resolvedHeader = false;
            for (var row = 0; row < sheet.Table.Rows.Count; row++)
            {
                var values = sheet.Table.Rows[row].ItemArray;
                if (!resolvedHeader)
                {
                    if (IsEmpty(values)) continue;
                    resolvedHeader = true;
                    LoadHeaders(table, values);
                }
                else
                {
                    LoadRow(table, values);
                }
            }

            if (dropBlank)
            {
                DropBlankRowsAndColumns(table);
            }

            return table;
        }

        private static void DropBlankRowsAndColumns(NExcelTable table)
        {
            table.Rows.RemoveAll(i => i.IsBlank);
            var blanks = table.Columns.Where(i => i.IsBlank).Select(i => i.Name).ToArray();
            table.Columns.RemoveAll(i => i.IsBlank);

            // Patch
            foreach (var row in table.Rows)
            {
                foreach (var blank in blanks)
                {
                    row.Values?.Remove(blank);
                }
            }
        }

        private static void LoadRow(NExcelTable table, object[] values)
        {
            var row = new NExcelRow();
            if (IsEmpty(values))
            {
                row.IsBlank = true;
                table.Rows.Add(row);
                return;
            }

            row.IsBlank = true;
            row.Values = new Dictionary<string, object>();
            for (var i = 0; i < values.Length; i++)
            {
                var column = table.Columns[i];
                var value = GetValue(values, i);
                row.Values[column.Name] = value;
                if (value != null)
                {
                    row.IsBlank = false;
                }

                if (value != null && column.IsBlank)
                {
                    column.IsBlank = false;
                    table.Columns[i] = column;
                }
            }

            table.Rows.Add(row);
        }

        private static void LoadHeaders(NExcelTable table, object[] values)
        {
            for (var i = 0; i < values.Length; i++)
            {
                var value = GetValue(values, i);
                var name = value == null ? $"Auto-{Guid.NewGuid()}" : $"{value}";
                if (table.Columns.Any(j => j.Name == name))
                {
                    name = $"{name}-{Guid.NewGuid()}";
                }

                table.Columns.Add(new NExcelColumn()
                {
                    Index = i,
                    IsBlank = true,
                    Name = name,
                });
            }
        }

        private static object GetValue(object[] values, int i)
        {
            var value = values[i];
            if (value == null) return null;
            if (value is DBNull) return null;
            if (!(value is string strValue)) return value;
            return string.IsNullOrEmpty(strValue) ? null : value;
        }

        private static bool IsEmpty(object[] values)
        {
            return values.Select((t, i) => GetValue(values, i)).All(value => value == null);
        }

        public static NExcelSimpleTable AsSimpleTable(NExcelTable table)
        {
            var rows = table.Rows.Select(RowAsList).Select(simpleRow => string.Join(", ", simpleRow.ToArray())).ToArray();
            return new NExcelSimpleTable()
            {
                rows = rows
            };
        }

        private static List<string> RowAsList(NExcelRow row)
        {
            var values = new List<string>();
            if (row.IsBlank || row.Values == null)
            {
                return values;
            }

            foreach (var key in row.Values.Keys)
            {
                values.Add(row.Values != null ? $"{key}: {row.Values[key]}" : $"{key}: NULL");
            }

            return values;
        }
    }
}