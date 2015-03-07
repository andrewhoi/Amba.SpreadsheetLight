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
    /// Chart customization options for bar and column charts.
    /// </summary>
    public class SLBarChartOptions
    {
        internal ushort iGapWidth;
        /// <summary>
        /// The gap width between bar or column clusters (in the same data series) as a percentage of bar or column width, ranging between 0% and 500% (both inclusive). The default is 150%.
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

        internal ushort iGapDepth;
        /// <summary>
        /// The gap depth between bar or columns clusters (between different data series) as a percentage of bar or column width, ranging between 0% and 500% (both inclusive). The default is 150%. This is only used for 3D chart version.
        /// </summary>
        public ushort GapDepth
        {
            get { return iGapDepth; }
            set
            {
                iGapDepth = value;
                if (iGapDepth > 500) iGapDepth = 500;
            }
        }

        internal sbyte byOverlap;
        /// <summary>
        /// The amount of overlapping for bars and columns on 2D bar/column charts, ranging from -100 to 100 (both inclusive). The default is 0. For stacked and "100% stacked" bar/column charts, this should be 100.
        /// </summary>
        public sbyte Overlap
        {
            get { return byOverlap; }
            set
            {
                byOverlap = value;
                if (byOverlap < -100) byOverlap = -100;
                if (byOverlap > 100) byOverlap = 100;
            }
        }

        /// <summary>
        /// Initializes an instance of SLBarChartOptions.
        /// </summary>
        public SLBarChartOptions()
        {
            this.iGapWidth = 150;
            this.iGapDepth = 150;
            this.byOverlap = 0;
        }

        internal SLBarChartOptions Clone()
        {
            SLBarChartOptions bco = new SLBarChartOptions();
            bco.iGapWidth = this.iGapWidth;
            bco.iGapDepth = this.iGapDepth;
            bco.byOverlap = this.byOverlap;

            return bco;
        }
    }
}
