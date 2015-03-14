// The MIT License (MIT)
//
// Copyright (c) 2015 Vladimir Kuznetsov
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

namespace Amba.SpreadsheetLight.Test
{
    using System.IO;
    using Xunit;
    using System.Linq;

    public class SLDocumentExt
    {
        private static string testTemplatesPath = "Amba.SpreadsheetLight.Test.Files";

        [Theory]
        [InlineData("A1", 1, 1, true)]
        [InlineData("Sheet!A1", 1, 1, true)]
        [InlineData("Sheet!A1:A2", 1, 2, true)]
        [InlineData("$A$1", 1, 1, true)]
        [InlineData("Sheet!$A$1", 1, 1, true)]
        [InlineData("Sheet!$A$1:$A$2", 1, 2, true)]
        [InlineData("Sheet!8:8", 8, 1, true)]
        [InlineData("Sheet!$8:$8", 8, 1, true)]
        [InlineData("Sheet!$8:$10", 8, 3, true)]
        [InlineData("Sheet!A:A", 0, 0, false)]
        [InlineData("Sheet!A:B", 0, 0, false)]
        public void WhatIsRowStartRowCount(string range, int row, int rowCount, bool result)
        {
            int r, n; r = n = 0;
            var actual = SLDocument.WhatIsRowStartRowCount(range, out r, out n);
            Assert.Equal(row, r);
            Assert.Equal(rowCount, n);
            Assert.Equal(result, actual);
        }

        [Fact]
        public void ColumnWidth()
        {
            // Ensure rowheight after inserting, deleting
            // Arrange
            var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()) + ".xlsx";
            File.Copy(Path.Combine(testTemplatesPath, "InsertingDeleting.xlsx"), tempFile);
            using (var doc = new SLDocument(tempFile))
            {
                // Act
                if (doc.SelectWorksheet("Sheet6"))
                {
                    doc.InsertColumn(1, 1);
                    doc.SetCellValue("B1", "I am inserted column, my width is 50.");
                    doc.SetColumnWidth(2, 50);
                }
                // Assert
                Assert.True(doc.GetColumnWidth(3) > 60); // Find out what wrong with this sizes during loading document
                Assert.True(doc.GetColumnWidth(4) < doc.GetColumnWidth(5));

                // Act
                if (doc.SelectWorksheet("Sheet6"))
                {
                    doc.DeleteColumn(1, 1);
                }
                // Assert
                Assert.True(doc.GetColumnWidth(2) > 60); // Find out what wrong with this sizes during loading document
                Assert.True(doc.GetColumnWidth(3) < doc.GetColumnWidth(4));
                doc.Save();
            }
            //System.Diagnostics.Process.Start(tempFile);
            File.Delete(tempFile);
            Assert.False(File.Exists(tempFile));
        }

        [Fact]
        public void RowHeight()
        {
            // Ensure rowheight after inserting, deleting
            // Arrange
            var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()) + ".xlsx";
            File.Copy(Path.Combine(testTemplatesPath, "InsertingDeleting.xlsx"), tempFile);
            using (var doc = new SLDocument(tempFile))
            {
                // Act
                if (doc.SelectWorksheet("Sheet5"))
                {
                    doc.InsertRow(1, 1);
                    doc.SetCellValue("A1", "I am inserted row, my height is 45.");
                    doc.SetRowHeight(1, 45);
                }
                // Assert
                Assert.Equal(30, doc.GetRowHeight(4));
                Assert.Equal(25.5, doc.GetRowHeight(5)); // TODO find out why whet set 25 here will be 25.5
                Assert.Equal(50.25, doc.GetRowHeight(6));

                // Act
                if (doc.SelectWorksheet("Sheet5"))
                {
                    doc.DeleteRow(1, 2);
                }
                // Assert
                Assert.Equal(30, doc.GetRowHeight(2));
                Assert.Equal(25.5, doc.GetRowHeight(3));
                Assert.Equal(50.25, doc.GetRowHeight(4));

                doc.Save();
            }
            //System.Diagnostics.Process.Start(tempFile);
            File.Delete(tempFile);
            Assert.False(File.Exists(tempFile));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void CopyColumn(bool toCut)
        {
            // Arrange
            var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()) + ".xlsx";
            File.Copy(Path.Combine(testTemplatesPath, "InsertingDeleting.xlsx"), tempFile);
            // Act
            using (var doc = new SLDocument(tempFile))
            {
                if (doc.SelectWorksheet("Sheet4"))
                {
                    doc.CopyColumn(1, 2, 3, toCut);
                }

                // Assert
                var cellC4 = doc.GetCells().Where(c => c.Key.RowIndex == 4 && c.Key.ColumnIndex == 3).FirstOrDefault();
                var cellD4 = doc.GetCells().Where(c => c.Key.RowIndex == 4 && c.Key.ColumnIndex == 4).FirstOrDefault();
                Assert.Equal("SUM(C1:C3)", cellC4.Value.CellFormula.FormulaText);
                Assert.Equal("SUM(D1:D3)", cellD4.Value.CellFormula.FormulaText);

                doc.Save();
            }
            //System.Diagnostics.Process.Start(tempFile);
            File.Delete(tempFile);
            Assert.False(File.Exists(tempFile));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void InsertCopyRow(bool toCut)
        {
            // Arrange
            var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()) + ".xlsx";
            File.Copy(Path.Combine(testTemplatesPath, "InsertingDeleting.xlsx"), tempFile);
            // Act
            using (var doc = new SLDocument(tempFile))
            {
                if (doc.SelectWorksheet("Sheet1"))
                {
                    // insert before range to move
                    doc.InsertRow(1, 1);
                    doc.InsertColumn(1, 1);
                    // delete inside range
                    doc.DeleteRow(4, 1);
                    doc.DeleteColumn(3, 1);
                    // Assert
                    Assert.Equal("Sheet1!$B$3:$C$4", doc.GetDefinedNameText("NamedRange"));
                }
                if (doc.SelectWorksheet("Sheet2"))
                {
                    // delete inside range
                    doc.DeleteRow(3, 1);
                    doc.DeleteColumn(2, 1);

                    doc.DeleteRow(7, 1);
                }
                if (doc.SelectWorksheet("Sheet3"))
                {
                    doc.CopyRow(2, 3, 4, toCut);
                    // C4
                    var cellC4 = doc.GetCells().Where(c => c.Key.RowIndex == 4 && c.Key.ColumnIndex == 3).FirstOrDefault();
                    // C5
                    var cellC5 = doc.GetCells().Where(c => c.Key.RowIndex == 5 && c.Key.ColumnIndex == 3).FirstOrDefault();

                    Assert.Equal("SUM(A4:B4)", cellC4.Value.CellFormula.FormulaText);
                    Assert.Equal("SUM(A5:B5)", cellC5.Value.CellFormula.FormulaText);
                }


                doc.SaveAs(tempFile);
            }
            //System.Diagnostics.Process.Start(tempFile);
            File.Delete(tempFile);
            Assert.False(File.Exists(tempFile));
        }

        [Fact]
        public void SelectWorksheet()
        {
            // Arrange
            var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()) + ".xlsx";
            File.Copy(Path.Combine(testTemplatesPath, "SetRangeValues.xlsx"), tempFile);
            // Act
            using (var doc = new SLDocument(tempFile))
            {
                if (doc.SelectWorksheet("Long name with space"))
                {
                    doc.SetCellValue("A1", "Hello, world!");
                }
                if (doc.SelectWorksheet("Длинное имя с пробелами"))
                {
                    doc.SetCellValue("A1", "Привет, мир!");
                }
                doc.SaveAs(tempFile);
            }
            // Assert
            using (var doc = new SLDocument(tempFile))
            {
                if (doc.SelectWorksheet("Long name with space"))
                {
                    Assert.Equal("Hello, world!", doc.GetCellValueAsString("A1"));
                }
                if (doc.SelectWorksheet("Длинное имя с пробелами"))
                {
                    Assert.Equal("Привет, мир!", doc.GetCellValueAsString("A1"));
                }
                doc.CloseWithoutSaving();
            }
            //System.Diagnostics.Process.Start(tempFile);
            File.Delete(tempFile);
            Assert.False(File.Exists(tempFile));
        }

        [Fact]
        public void SetDefinedNameValue()
        {
            // Arrange
            var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()) + ".xlsx";
            File.Copy(Path.Combine(testTemplatesPath, "SetRangeValues.xlsx"), tempFile);
            // Act
            using (var doc = new SLDocument(tempFile))
            {
                doc.SetDefinedNameValue<string>("SimpleNamedRange", "I am SimpleNamedRange");
                doc.SetDefinedNameValue<string>("MergedNamedRange", "I am MergedNamedRange");

                doc.SetDefinedNameValue<string>("SeveralCells", "Several cells value");
                doc.SetDefinedNameValue<string>("SeveralRanges", "Several ranges value");
                doc.SetDefinedNameValue<string>("SeveralMergedRanges", "I am SeveralMergedRanges");

                //doc.SetRangeValue<string>("EntireColumn", "I am EntireColumn"); // very slowly on saving file
                //// Assert
                //Assert.Equal("I am EntireColumn", doc.GetCellValueAsString("H1"));
                //Assert.Equal("I am EntireColumn", doc.GetCellValueAsString("H100"));
                //Assert.Equal("I am EntireColumn", doc.GetCellValueAsString("H1000000"));
                //doc.DeleteColumn("H", 1);

                doc.SetDefinedNameValue<string>("EntireRow", "EntireRow value");

                doc.Save();
            }

            // Assert
            using (var doc = new SLDocument(tempFile))
            {
                Assert.Equal("I am SimpleNamedRange", doc.GetCellValueAsString("B1"));
                Assert.Equal("I am MergedNamedRange", doc.GetCellValueAsString("B2"));
                Assert.Equal("EntireRow value", doc.GetCellValueAsString("A5"));
                Assert.Equal("EntireRow value", doc.GetCellValueAsString("XFD5"));
                Assert.Equal("Several cells value", doc.GetCellValueAsString("B7"));
                Assert.Equal("Several cells value", doc.GetCellValueAsString("D7"));
                Assert.Equal("Several cells value", doc.GetCellValueAsString("F7"));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString("B9"));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString("B10"));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString("B11"));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString("D9"));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString("D10"));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString("D11"));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString("F9"));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString("F10"));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString("F11"));
                Assert.Equal("I am SeveralMergedRanges", doc.GetCellValueAsString("B13"));
                Assert.Equal("I am SeveralMergedRanges", doc.GetCellValueAsString("D13"));
                Assert.Equal("I am SeveralMergedRanges", doc.GetCellValueAsString("F13"));
                doc.CloseWithoutSaving();
            }
            //System.Diagnostics.Process.Start(tempFile);
            File.Delete(tempFile);
            Assert.False(File.Exists(tempFile));
        }

        [Theory]
        [InlineData(1, 0)]
        [InlineData(-1, 0)]
        [InlineData(0, 1)]
        [InlineData(0, -1)]
        [InlineData(-1, -1)]
        [InlineData(1, 1)]
        public void SetDefinedNameValueWithOffsets(int rowOffset, int columnOffset)
        {
            // Arrange
            var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()) + ".xlsx";
            File.Copy(Path.Combine(testTemplatesPath, "SetRangeValues.xlsx"), tempFile);

            // Act
            using (var doc = new SLDocument(tempFile))
            {
                doc.SetDefinedNameValue<string>("SimpleNamedRange", "I am SimpleNamedRange", rowOffset, columnOffset);
                doc.SetDefinedNameValue<string>("MergedNamedRange", "I am MergedNamedRange", rowOffset, columnOffset);

                doc.SetDefinedNameValue<string>("SeveralCells", "Several cells value", rowOffset, columnOffset);
                doc.SetDefinedNameValue<string>("SeveralRanges", "Several ranges value", rowOffset, columnOffset);
                doc.SetDefinedNameValue<string>("SeveralMergedRanges", "I am SeveralMergedRanges", rowOffset, columnOffset);
                doc.SetDefinedNameValue<string>("EntireRow", "EntireRow value", rowOffset, columnOffset);
                doc.Save();
            }

            // Assert
            using (var doc = new SLDocument(tempFile))
            {
                if (rowOffset > 0)
                {
                    Assert.Equal("I am SimpleNamedRange", doc.GetCellValueAsString(1 + rowOffset, 2 + columnOffset));
                }
                Assert.Equal("I am MergedNamedRange", doc.GetCellValueAsString(2 + rowOffset, 2 + columnOffset));
                Assert.Equal("EntireRow value", doc.GetCellValueAsString(5 + rowOffset, 2 + columnOffset));
                Assert.Equal("EntireRow value", doc.GetCellValueAsString(5 + rowOffset, 1000 + columnOffset));
                Assert.Equal("Several cells value", doc.GetCellValueAsString(7 + rowOffset, 2 + columnOffset));
                Assert.Equal("Several cells value", doc.GetCellValueAsString(7 + rowOffset, 4 + columnOffset));
                Assert.Equal("Several cells value", doc.GetCellValueAsString(7 + rowOffset, 6 + columnOffset));

                Assert.Equal("Several ranges value", doc.GetCellValueAsString(9 + rowOffset, 2 + columnOffset));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString(10 + rowOffset, 2 + columnOffset));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString(11 + rowOffset, 2 + columnOffset));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString(9 + rowOffset, 4 + columnOffset));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString(10 + rowOffset, 4 + columnOffset));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString(11 + rowOffset, 4 + columnOffset));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString(9 + rowOffset, 6 + columnOffset));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString(10 + rowOffset, 6 + columnOffset));
                Assert.Equal("Several ranges value", doc.GetCellValueAsString(11 + rowOffset, 6 + columnOffset));

                Assert.Equal("I am SeveralMergedRanges", doc.GetCellValueAsString(13 + rowOffset, 2 + columnOffset));
                Assert.Equal("I am SeveralMergedRanges", doc.GetCellValueAsString(13 + rowOffset, 4 + columnOffset));
                Assert.Equal("I am SeveralMergedRanges", doc.GetCellValueAsString(13 + rowOffset, 6 + columnOffset));
                doc.CloseWithoutSaving();
            }
            //System.Diagnostics.Process.Start(tempFile);
            File.Delete(tempFile);
            Assert.False(File.Exists(tempFile));
        }
    }
}
