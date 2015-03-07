// Copyright (c) 2011 Vincent Tan Wai Lip
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
using System;

namespace Amba.SpreadsheetLight.Charts
{
    /// <summary>
    /// General chart customization options on chart creation.
    /// </summary>
    public class SLCreateChartOptions
    {
        /// <summary>
        /// True if rows contain the data series. False if columns contain the data series.
        /// Set to null if SpreadsheetLight is to determine data series orientation.
        /// If the number of columns in a given cell range is more than or equal to the
        /// number of rows, then it's decided that rows contain data series (else it's columns).
        /// The default value is null.
        /// </summary>
        public bool? RowsAsDataSeries { get; set; }

        /// <summary>
        /// True if hidden data is used in the chart. False otherwise.
        /// The default value is false.
        /// </summary>
        public bool ShowHiddenData { get; set; }

        /// <summary>
        /// True to use default chart styling from latest version of Excel
        /// (but no guarantees on conformity or Excel version). False otherwise.
        /// The default value is true.
        /// </summary>
        public bool IsStylish { get; set; }

        /// <summary>
        /// Initializes an instance of SLCreateChartOptions.
        /// </summary>
        public SLCreateChartOptions()
        {
            this.RowsAsDataSeries = null;
            this.ShowHiddenData = false;
            this.IsStylish = true;
        }
    }
}
