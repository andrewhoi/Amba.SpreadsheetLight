## Amba.SpreadsheetLight

The library is for reading and creating Excel files in OpenXml format.
Compatible wtih Microsoft Excel 2007/2010/2013 and LibreOffice Calc.

This is the fork from www.spreadsheetlight.com library with some improvements.

Library was merged with required old version of DocumentFormat.OpenXml to allow use new one.



####How to start
    
Install with nuget:

    PM> Install-Package Amba.SpreadsheetLight

Example:

    using (SLDocument sl = new SLDocument("Template.xlsx"))
    {
        if (sl.SelectWorksheet("Sheet1"))
        {
            sl.SetCellValue("A1", "Hello, world!");
        }
        sl.SaveAs("GeneratedReport.xlsx");
    }

You can find more tutorial and examples on [Tutorial](http://spreadsheetlight.com/tutorial/) and [Samples](http://spreadsheetlight.com/sample-code/) on [SpreadsheetLight site](www.spreadsheetlight.com).

## Changelog

**1.0.0** (March, 05 2015)

* Added: **SetDefinedNameValue&lt;T&gt;** method to SLDocument class.
* Fixed: restore **row heights** after inserting/deleting row(s).
* dll merged with required old version of DocumentFormat.OpenXml to allow use new one.

## License

The MIT License (MIT)

Copyright (c) 2011-2015 Vladimir Kuznetsov, Vincent Tan Wai Lip

