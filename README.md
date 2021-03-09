# Excel reader

Provides a basic excel format reader for unity.

This is just a straight out wrapper of the https://github.com/ExcelDataReader/ExcelDataReader
with the most basic possible wrapper to allow it be used in unity.

## Usage

See the tests in the `tests` folder for each class for usage examples.

Basic usage at runtime:

```
public TextAsset asset;

public void ReadFile()
{
    var document = new NExcel().Read(asset);
    var table = document.Sheets[0].LoadTable();
    Debug.Log(table.AsJson());
}
```

In editor scripts there's no need to rename the file and load it as a text asset; just use:

```
public void TestCanReadTableComplex()
{
    var document = new NExcel().Read(Path.Combine(Here(), "ExcelTestsDocument.xlsx"));
    var sheet = document.Sheets.FirstOrDefault(i => i.Name == "Mancy");
    var table = sheet.LoadTable();
    Debug.Log(table.AsJson());
}

private static string Here([CallerFilePath] string path = null)
{
    return Path.GetDirectoryName(path);
}
```

Note that the `LoadTable` function is an opinionated helper that loads an entire sheet
as a table; it probably isn't suitable for most non-trivial purposes; just use the DataTable
object on the sheet.

## Install

From your unity project folder:

    npm init
    npm install TEMPLATE --save
    echo Assets/pkg-all >> .gitignore
    echo Assets/pkg-all.meta >> .gitignore

The package and all its dependencies will be installed in
your Assets/pkg-all folder.

## Development

    cd test
    npm install

To reinstall the files from the src folder, run `npm install ..` again.

### Tests

All tests are located in the `Tests` folder.
