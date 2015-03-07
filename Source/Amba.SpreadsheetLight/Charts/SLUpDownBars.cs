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
using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Amba.SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for up-down bars.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Charts.UpDownBars class.
    /// </summary>
    public class SLUpDownBars
    {
        internal ushort iGapWidth;
        /// <summary>
        /// The gap width between consecutive up-down bars as a percentage of the width of the bar, ranging from 0 to 500 (both inclusive).
        /// </summary>
        public ushort GapWidth
        {
            get { return iGapWidth; }
            set
            {
                iGapWidth = value;
                if (iGapWidth > 500) iGapWidth = 500;
            }
        }

        /// <summary>
        /// The up bars.
        /// </summary>
        public SLUpBars UpBars { get; set; }

        /// <summary>
        /// The down bars.
        /// </summary>
        public SLDownBars DownBars { get; set; }

        internal SLUpDownBars(List<System.Drawing.Color> ThemeColors, bool IsStylish = false)
        {
            this.iGapWidth = 150;
            this.UpBars = new SLUpBars(ThemeColors, IsStylish);
            this.DownBars = new SLDownBars(ThemeColors, IsStylish);
        }

        internal C.UpDownBars ToUpDownBars(bool IsStylish = false)
        {
            C.UpDownBars udb = new C.UpDownBars();
            udb.GapWidth = new C.GapWidth() { Val = iGapWidth };
            udb.UpBars = this.UpBars.ToUpBars(IsStylish);
            udb.DownBars = this.DownBars.ToDownBars(IsStylish);

            return udb;
        }

        internal SLUpDownBars Clone()
        {
            SLUpDownBars udb = new SLUpDownBars(new List<System.Drawing.Color>());
            udb.iGapWidth = this.iGapWidth;
            udb.UpBars = this.UpBars.Clone();
            udb.DownBars = this.DownBars.Clone();

            return udb;
        }
    }
}
