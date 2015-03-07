## Amba.SpreadsheetLight

This is the fork from www.spreadsheetlight.com library.

This library is for reading and creating Excel (LibreOffice Calc) files in OpenXml format and based on "DocumentFormat.OpenXml 2.0.0" library.

####How to start
    
    using (SLDocument sl = new SLDocument("Template.xlsx"))
    {
        if (sl.SelectWorksheet("Sheet1"))
        {
            sl.SetCellValue("A1", "Hello, world!");
        }
        sl.SaveAs("GeneratedReport.xlsx");
    }

You can find more tutorial and examples on [Tutorial](http://spreadsheetlight.com/tutorial/) and [Samples](http://spreadsheetlight.com/sample-code/) on [SpreadsheetLight site](www.spreadsheetlight.com).

