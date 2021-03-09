using System.Collections;
using System.Collections.Generic;
using N.Package.Excel;
using N.Package.Excel.Model;
using N.Package.Tests.Runtime;
using UnityEngine;

public class ExcelRuntimeTests : RuntimeTest
{
    public TextAsset asset;

    [RuntimeTest]
    public void TestReadAsset()
    {
        var document = new NExcel().Read(asset);
        var table = document.Sheets[0].LoadTable();
        Debug.Log(table.AsJson());
        Completed();
    }
}