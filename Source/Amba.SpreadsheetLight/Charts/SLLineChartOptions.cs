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
    /// Chart customization options for line charts.
    /// </summary>
    public class SLLineChartOptions
    {
        internal ushort iGapDepth;
        /// <summary>
        /// The gap depth between line clusters (between different data series) as a percentage of bar or column width, ranging between 0% and 500% (both inclusive). The default is 150%. This is only used for 3D chart version.
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
        /// Indicates if the line chart has drop lines.
        /// </summary>
        public bool HasDropLines { get; set; }

        /// <summary>
        /// Drop lines properties.
        /// </summary>
        public SLDropLines DropLines { get; set; }

        /// <summary>
        /// Indicates if the line chart has high-low lines. This is not applicable for 3D line charts.
        /// </summary>
        public bool HasHighLowLines { get; set; }

        /// <summary>
        /// High-low lines properties.
        /// </summary>
        public SLHighLowLines HighLowLines { get; set; }

        /// <summary>
        /// Indicates if the line chart has up-down bars. This is not applicable for 3D line charts.
        /// </summary>
        public bool HasUpDownBars { get; set; }

        /// <summary>
        /// Up-down bars properties.
        /// </summary>
        public SLUpDownBars UpDownBars { get; set; }

        /// <summary>
        /// Whether the line connecting data points use C splines (instead of straight lines).
        /// </summary>
        public bool Smooth { get; set; }

        /// <summary>
        /// Initializes an instance of SLLineChartOptions. It is recommended to use SLChart.CreateLineChartOptions().
        /// </summary>
        public SLLineChartOptions()
        {
            this.Initialize(new List<System.Drawing.Color>(), false);
        }

        internal SLLineChartOptions(List<System.Drawing.Color> ThemeColors, bool IsStylish = false)
        {
            this.Initialize(ThemeColors, IsStylish);
        }

        private void Initialize(List<System.Drawing.Color> ThemeColors, bool IsStylish)
        {
            this.iGapDepth = 150;
            this.HasDropLines = false;
            this.DropLines = new SLDropLines(ThemeColors, IsStylish);
            this.HasHighLowLines = false;
            this.HighLowLines = new SLHighLowLines(ThemeColors, IsStylish);
            this.HasUpDownBars = false;
            this.UpDownBars = new SLUpDownBars(ThemeColors, IsStylish);
            this.Smooth = false;
        }

        /// <summary>
        /// Clone an instance of SLLineChartOptions.
        /// </summary>
        /// <returns>An SLLineChartOptions object.</returns>
        public SLLineChartOptions Clone()
        {
            SLLineChartOptions lco = new SLLineChartOptions();
            lco.iGapDepth = this.iGapDepth;
            lco.HasDropLines = this.HasDropLines;
            lco.DropLines = this.DropLines.Clone();
            lco.HasHighLowLines = this.HasHighLowLines;
            lco.HighLowLines = this.HighLowLines.Clone();
            lco.HasUpDownBars = this.HasUpDownBars;
            lco.UpDownBars = this.UpDownBars.Clone();
            lco.Smooth = this.Smooth;

            return lco;
        }
    }
}
