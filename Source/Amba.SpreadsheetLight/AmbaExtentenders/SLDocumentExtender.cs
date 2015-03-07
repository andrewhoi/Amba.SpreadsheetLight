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

namespace Amba.SpreadsheetLight
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Extensions for SLDocument class
    /// </summary>
    public static class SLDocumentExtender
    {
        /// <summary>
        /// Set value to DefinedName range
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="doc"></param>
        /// <param name="definedName"></param>
        /// <param name="value"></param>
        public static void SetDefinedNameValue<T>(this SLDocument doc, string definedName, T value)
        {
            doc.SetDefinedNameValue<T>(definedName, value, 0, 0);
        }

        /// <summary>
        /// Set value to DefinedName range with offset
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="doc"></param>
        /// <param name="definedName"></param>
        /// <param name="value"></param>
        /// <param name="rowOffset"></param>
        /// <param name="columnOffset"></param>
        public static void SetDefinedNameValue<T>(this SLDocument doc, string definedName, T value, int rowOffset = 0, int columnOffset = 0)
        {
            var address = doc.GetDefinedNameText(definedName);
            if (String.IsNullOrWhiteSpace(address)) return;

            var @switch = new Dictionary<Type, Action<int, int, object>> {
                { typeof(string), (row, col, val) => { doc.SetCellValue(row, col, (string)val); } },
                { typeof(int), (row, col, val) => { doc.SetCellValue(row, col, (int)val); } },
                { typeof(float), (row, col, val) => { doc.SetCellValue(row, col, (float)val); } },
                { typeof(double), (row, col, val) => { doc.SetCellValue(row, col, (double)val); } },
                { typeof(decimal), (row, col, val) => { doc.SetCellValue(row, col, (decimal)val); } },
                { typeof(DateTime), (row, col, val) => { doc.SetCellValue(row, col, (DateTime)val); } },
                { typeof(bool), (row, col, val) => { doc.SetCellValue(row, col, (bool)val); } },
            };

            var addresses = address.Split(',');
            for (int i = 0; i < addresses.Length; i++)
            {
                var adr = addresses[i];
                // parse full address
                var indexExl = adr.IndexOf('!');
                if (indexExl == -1) continue;
                var sheetName = adr.Substring(0, indexExl);
                if (adr.Contains(':')) // range in address
                {
                    var addressRange = adr.Substring(indexExl + 1);
                    var rangeArray = addressRange.Split(':');
                    var beginArray = rangeArray[0].Split('$');
                    var endArray = rangeArray[1].Split('$');
                    // for cycle
                    int rowStart = -1;
                    int colStart = -1;
                    int rowEnd = -1;
                    int colEnd = -1;
                    bool doCycle = false;
                    if (beginArray.Length == 2) // entire column or row
                    {
                        int val = -1;
                        if (Int32.TryParse(beginArray[1], out val)) // is entire row
                        {
                            rowStart = val;
                            doCycle = Int32.TryParse(endArray[1], out rowEnd);
                            colStart = 1;
                            colEnd = SLDocument.ColumnLimit;

                            rowStart += rowOffset;
                            rowEnd += rowOffset;
                            if (rowStart <= 0)
                            {
                                if (rowEnd <= 0) doCycle = false;
                                else // shrink range
                                    rowStart = 1;
                            }
                        }
                        else // entire column
                        {
                            rowStart = 1;
                            rowEnd = SLDocument.RowLimit; // very slowly!!!! Don't do it
                            colStart = SLConvert.ToColumnIndex(beginArray[1]);
                            colEnd = SLConvert.ToColumnIndex(endArray[1]);
                            colStart += columnOffset;
                            colEnd += columnOffset;
                            doCycle = colStart > 0 && colEnd > 0;
                        }
                    }
                    else // simple range 
                    {
                        var correctStartCell = SLDocument.WhatIsRowColumnIndex(rangeArray[0].Replace("$", "")
                            , out rowStart, out colStart);
                        var correctEndCell = SLDocument.WhatIsRowColumnIndex(rangeArray[1].Replace("$", "")
                            , out rowEnd, out colEnd);
                        doCycle = correctStartCell && correctEndCell;
                        // do with offsets
                        rowStart += rowOffset;
                        rowEnd += rowOffset;
                        if (rowStart <= 0)
                        {
                            if (rowEnd <= 0) doCycle = false;
                            else // shrink range vertically
                                rowStart = 1;
                        }
                        colStart += columnOffset;
                        colEnd += columnOffset;
                        if (colStart <= 0)
                        {
                            if (colEnd <= 0) doCycle = false;
                            else // shrink range horizontally
                                colStart = 1;
                        }
                    }


                    // Fill cells
                    if (doCycle)
                    {
                        // select worksheet and insert value
                        if (doc.SelectWorksheet(sheetName))
                        {
                            for (int r = rowStart; r <= rowEnd; r++)
                            {
                                for (int c = colStart; c <= colEnd; c++)
                                {
                                    @switch[typeof(T)](r, c, value);
                                }
                            }
                        }
                    }
                }
                else // single cells in address
                {

                    var addressCell = adr.Substring(indexExl + 1).Replace("$", "");
                    int row = -1;
                    int col = -1;
                    if (SLDocument.WhatIsRowColumnIndex(addressCell, out row, out col))
                    {
                        row += rowOffset;
                        col += columnOffset;
                        // select worksheet and insert value
                        if (row > 0 && row <= SLDocument.RowLimit
                            && col > 0 && col <= SLDocument.ColumnLimit
                            && doc.SelectWorksheet(sheetName))
                        {
                            @switch[typeof(T)](row, col, value);
                        }
                    }
                }
            }
        } // SetRangeValue<T>

        /// <summary>
        /// Save rows' heights before deleting/inserting
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="rowStart"></param>
        /// <returns></returns>
        internal static double[] GetRowHeights(this SLDocument doc, int rowStart)
        {
            var stat = doc.GetWorksheetStatistics();
            int rowCount = stat.EndRowIndex - rowStart + 1;
            rowCount = rowCount < 0 ? 0 : rowCount;
            var result = new double[rowCount];
            for (int i = 0; i < rowCount; i++)
            {
                result[i] = doc.GetRowHeight(i + rowStart);
            }
            return result;
        }
        /// <summary>
        /// Restore rows' heights after inserting, deleting
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="startRow"></param>
        /// <param name="heights"></param>
        internal static void RestoreRowHeight(this SLDocument doc, int startRow, double[] heights)
        {
            for (int i = 0; i < heights.Length; i++)
            {
                doc.SetRowHeight(startRow + i, heights[i]);
            }
        }
    }
}
