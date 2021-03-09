using System.Collections.Generic;
using System.Data;
using System.IO;
using ExcelDataReader;
using N.Package.Excel.Model;
using UnityEditor.VersionControl;
using UnityEngine;
using FileMode = System.IO.FileMode;

namespace N.Package.Excel
{
    public class NExcel
    {
        public NExcelDocument Read(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    return new NExcelDocument(result);
                }
            }
        }

        public NExcelDocument Read(TextAsset asset)
        {
            using (var stream = new MemoryStream(asset.bytes))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    return new NExcelDocument(result);
                }
            }
        }
    }
}