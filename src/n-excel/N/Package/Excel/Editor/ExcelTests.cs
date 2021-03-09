using System.IO;
using System.Linq;
using NUnit.Framework;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using N.Package.Excel;
using UnityEngine;

namespace Tests
{
    public class ExcelTests
    {
        [Test]
        public void TestCanFindFiles()
        {
            var document = Path.Combine(Here(), "ExcelTestsDocument.xlsx");
            Assert.IsTrue(File.Exists(document));
        }

        [Test]
        public void TestCanReadSheets()
        {
            var document = new NExcel().Read(Path.Combine(Here(), "ExcelTestsDocument.xlsx"));
            Assert.AreEqual(5, document.Sheets.Length);
            Assert.AreEqual("Mancy", document.Sheets[0].Name);
            Assert.AreEqual("Sheet2", document.Sheets[1].Name);
            Assert.AreEqual("Sheet 1", document.Sheets[2].Name);
            Assert.AreEqual("Sheet3", document.Sheets[3].Name);
            Assert.AreEqual("Hello", document.Sheets[4].Name);
        }

        [Test]
        public void TestCanReadTable()
        {
            var document = new NExcel().Read(Path.Combine(Here(), "ExcelTestsDocument.xlsx"));
            var sheet = document.Sheets.FirstOrDefault(i => i.Name == "Sheet2");
            Assert.IsNotNull(sheet);

            var table = sheet.LoadTable();
            Assert.IsTrue(table.Columns.Any(c => c.Name == "A"));
        }

        [Test]
        public void TestCanReadTableComplex()
        {
            var document = new NExcel().Read(Path.Combine(Here(), "ExcelTestsDocument.xlsx"));
            var sheet = document.Sheets.FirstOrDefault(i => i.Name == "Mancy");
            Assert.IsNotNull(sheet);

            var table = sheet.LoadTable();
            Assert.IsNotNull(table.AsJson());
            Debug.Log(table.AsJson());
        }

        private static string Here([CallerFilePath] string path = null)
        {
            return Path.GetDirectoryName(path);
        }
    }
}