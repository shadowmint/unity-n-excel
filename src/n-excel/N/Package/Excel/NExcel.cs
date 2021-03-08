using System.IO;
using ExcelDataReader;
using N.Package.Excel.Model;

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
    }
}