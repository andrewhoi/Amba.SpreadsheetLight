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

namespace Amba.SpreadsheetLight.Charts
{
    /// <summary>
    /// Chart customization options for area charts.
    /// </summary>
    public class SLAreaChartOptions
    {
        /// <summary>
        /// Indicates if the area chart has drop lines.
        /// </summary>
        public bool HasDropLines { get; set; }

        /// <summary>
        /// Drop lines properties.
        /// </summary>
        public SLDropLines DropLines { get; set; }

        internal ushort iGapDepth;
        /// <summary>
        /// The gap depth between area clusters (between different data series) as a percentage of width, ranging between 0% and 500% (both inclusive). The default is 150%. This is only used for 3D chart version.
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

        /// <summary>
        /// Initializes an instance of SLAreaChartOptions. It is recommended to use SLChart.CreateAreaChartOptions().
        /// </summary>
        public SLAreaChartOptions()
        {
            this.Initialize(new List<System.Drawing.Color>(), false);
        }

        internal SLAreaChartOptions(List<System.Drawing.Color> ThemeColors, bool IsStylish = false)
        {
            this.Initialize(ThemeColors, IsStylish);
        }

        private void Initialize(List<System.Drawing.Color> ThemeColors, bool IsStylish)
        {
            this.HasDropLines = false;
            this.DropLines = new SLDropLines(ThemeColors, IsStylish);
            this.iGapDepth = 150;
        }

        /// <summary>
        /// Clone a new instance of SLAreaChartOptions.
        /// </summary>
        /// <returns>An SLAreaChartOptions object.</returns>
        public SLAreaChartOptions Clone()
        {
            SLAreaChartOptions aco = new SLAreaChartOptions();
            aco.HasDropLines = this.HasDropLines;
            aco.DropLines = this.DropLines.Clone();
            aco.iGapDepth = this.iGapDepth;

            return aco;
        }
    }
}
