using System.Data;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using N.Package.Excel.Internal;

namespace N.Package.Excel.Model
{
    public class NExcelDocument
    {
        private readonly DataSet _result;

        public NExcelDocument(DataSet result)
        {
            _result = result;
            Sheets = NExcelUtils.LoadSheetsFrom(result);
        }

        public NExcelSheet[] Sheets { get; set; }
    }
}