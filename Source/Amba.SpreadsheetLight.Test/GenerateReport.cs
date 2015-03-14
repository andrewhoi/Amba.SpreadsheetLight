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

    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Xunit;
    using Newtonsoft.Json.Linq;

    public class GenerateReport
    {
        private static string testTemplatesPath = "Amba.SpreadsheetLight.Test.Files";

        /// <summary>
        /// Test must work correct with formulas on the inserted row
        /// </summary>
        [Fact]
        public void Test1()
        {
            // Arrange
            var tempFile = GetTempFileName();
            File.Copy(Path.Combine(testTemplatesPath, "ReportTemplate1.xlsx"), tempFile);
            using (var doc = new SLDocument(tempFile))
            {
                // Act
                foreach (var item in (JArray)jsonData1.GetValue("Items"))
                {
                    var rangeRow = doc.GetDefinedNameText("Items");
                    int r, cnt;
                    SLDocument.WhatIsRowStartRowCount(rangeRow, out r, out cnt);
                    doc.InsertRow(r + 1, cnt);
                    doc.CopyRow(r, r + cnt - 1, r + cnt);
                }
                doc.Save();
            }
            // Process.Start(tempFile);
            File.Delete(tempFile);
            Assert.False(File.Exists(tempFile));
        }
        private string GetTempFileName()
        {
            var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()) + ".xlsx";
            return tempFile;
        }
        private JObject jsonData1 = JObject.Parse(@"
{   'Num': '12345678-15',
    'Date': '2015-05-22T21:01:02',
    'Org': { 'Name': 'Customer\'s name here' },
    'Items': [
        { 'Item': 'Item #1', 'Qnt': 12.5, 'Price': 20.0 },
        { 'Item': 'Item #2', 'Qnt': 1, 'Price': 55.65 },
        { 'Item': 'Item #3', 'Qnt': 2.0, 'Price': 50.1 }
    ],
    'GroupRows': [{
        'Item' : 'Item\'s Group #1',
        'Details': [
            { 'Item': 'Item #1-1', 'Qnt': 11.5, 'Price': 21.0 },
            { 'Item': 'Item #1-2', 'Qnt': 10, 'Price': 45.99 },   
        ]},
        {
        'Item' : 'Item\'s Group #2',
        'Details': [
            { 'Item': 'Item #2-1', 'Qnt': 11.5, 'Price': 21.0 },
        ]}
    ]
}");
    }
}
