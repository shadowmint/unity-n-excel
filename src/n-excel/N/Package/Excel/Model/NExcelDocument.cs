using System.Data;

namespace N.Package.Excel.Model
{
    public class NExcelDocument
    {
        private readonly DataSet _result;

        public NExcelDocument(DataSet result)
        {
            _result = result;
        }

        public NExcelSheet[] Sheets { get; set; }
    }
}