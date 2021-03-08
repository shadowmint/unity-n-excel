using System.Collections;
using System.Collections.Generic;
using N.Package.Excel;
using N.Package.Excel.Model;
using UnityEngine;

public class ExcelRuntimeTests : MonoBehaviour
{
    public TextAsset asset;

    void Start()
    {
        var document = new NExcel().Read(asset);
        foreach (var sheet in document.Sheets ?? new NExcelSheet[] { })
        {
            Debug.Log(sheet.Name);
        }
    }
}