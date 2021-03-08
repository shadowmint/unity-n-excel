using System.IO;
using NUnit.Framework;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using UnityEngine;

namespace Tests
{
    public class ExcelTests
    {
        [Test]
        public void TestCanFindFiles()
        {
            var document = Path.Combine(Here(), "ExcelTestsDocument.xlsx");
            Debug.Log(document);
            Assert.IsTrue(File.Exists(document));
        }

        [Test]
        public void TestCanReadSheets()
        {
            var document = Path.Combine(Here(), "ExcelTestsDocument.xlsx");
        }

        private static string Here([CallerFilePath] string path = null)
        {
            return Path.GetDirectoryName(path);
        }
    }
}