# ExcelDataReader.ReadToClass

Uses https://github.com/ExcelDataReader/ExcelDataReader as Excel reading library and adds additional extensions to read sheets rows and cells into a specified class structure.

# Usage

There are two ways to use this library: 

## Using attributes
Define classes using attributes for sheets (`ExcelTableAttribute`) and columns (`ExcelColumnAttribute`):

```
public class OneSheetExcel
{
    [ExcelTable("My Sheet 1")]
    public List<FirstSheet> FirstSheetRows { get; set; }
}

public class FirstSheet
{
    [ExcelColumn("Text Column")]
    public string TextColumn { get; set; }

    [ExcelColumn("Some Int")]
    public int IntColumn { get; set; }
}
```
Call excelReader `AsClass` extension method without any additional configuration.
```
using (var fileStream = File.Open(excelFileName, FileMode.Open, FileAccess.Read))
using (var excelReader = ExcelReaderFactory.CreateReader(fileStream))
{
    var result = excelReader.AsClass<OneSheetExcel>();
}
```

## Using fluent configuration

Define classes without any attributes:
```
public class OneSheetExcel
{
    public List<FirstSheet> FirstSheetRows { get; set; }
}

public class FirstSheet
{
    public string TextColumn { get; set; }

    public int IntColumn { get; set; }
}
```
Use `FluentConfig.ConfigureFor` method to define relations between excel sheets and columns, and classes and properties.
Then call excelReader `AsClass` extension method with created configuration as parameter.
```
using (var fileStream = File.Open(excelFileName, FileMode.Open, FileAccess.Read))
using (var excelReader = ExcelReaderFactory.CreateReader(fileStream))
{
	var config = FluentConfig.ConfigureFor<OneSheetExcel>().WithTables(table =>
	{
		table.Bind("My Sheet 1", m => m.FirstSheetRows).WithColumns(column =>
		{
			column.Bind("Text Column", c => c.TextColumn);
			column.Bind("Some Int", c => c.IntColumn);
		});
	});

  var result = excelReader.AsClass<OneSheetExcel>(config);
}
```
Benefit of this approach is that you can use the same classes for different excel schemas by changing configuration only.
