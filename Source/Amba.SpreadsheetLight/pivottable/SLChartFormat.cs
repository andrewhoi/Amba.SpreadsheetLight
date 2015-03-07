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
using DocumentFormat.OpenXml.Spreadsheet;

namespace Amba.SpreadsheetLight
{
    internal class SLChartFormat
    {
        internal SLPivotArea PivotArea { get; set; }
        internal uint Chart { get; set; }
        internal uint Format { get; set; }
        internal bool Series { get; set; }

        internal SLChartFormat()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.PivotArea = new SLPivotArea();
            this.Chart = 0;
            this.Format = 0;
            this.Series = false;
        }

        internal void FromChartFormat(ChartFormat cf)
        {
            this.SetAllNull();

            if (cf.PivotArea != null) this.PivotArea.FromPivotArea(cf.PivotArea);

            if (cf.Chart != null) this.Chart = cf.Chart.Value;
            if (cf.Format != null) this.Format = cf.Format.Value;
            if (cf.Series != null) this.Series = cf.Series.Value;
        }

        internal ChartFormat ToChartFormat()
        {
            ChartFormat cf = new ChartFormat();
            cf.PivotArea = this.PivotArea.ToPivotArea();

            cf.Chart = this.Chart;
            cf.Format = this.Format;
            if (this.Series != false) cf.Series = this.Series;

            return cf;
        }

        internal SLChartFormat Clone()
        {
            SLChartFormat cf = new SLChartFormat();
            cf.PivotArea = this.PivotArea.Clone();
            cf.Chart = this.Chart;
            cf.Format = this.Format;
            cf.Series = this.Series;

            return cf;
        }
    }
}
