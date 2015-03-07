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
    internal class SLStringReference
    {
        internal string WorksheetName { get; set; }
        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        internal string Formula { get; set; }

        // this is StringCache
        internal uint PointCount { get; set; }
        /// <summary>
        /// This takes the place of StringCache
        /// </summary>
        internal List<SLStringPoint> Points { get; set; }

        internal SLStringReference()
        {
            this.WorksheetName = string.Empty;
            this.StartRowIndex = 1;
            this.StartColumnIndex = 1;
            this.EndRowIndex = 1;
            this.EndColumnIndex = 1;

            this.Formula = string.Empty;
            this.PointCount = 0;
            this.Points = new List<SLStringPoint>();
        }

        internal C.StringReference ToStringReference()
        {
            C.StringReference sr = new C.StringReference();
            sr.Formula = new C.Formula(this.Formula);
            sr.StringCache = new C.StringCache();
            sr.StringCache.PointCount = new C.PointCount() { Val = this.PointCount };
            for (int i = 0; i < this.Points.Count; ++i)
            {
                sr.StringCache.Append(this.Points[i].ToStringPoint());
            }

            return sr;
        }

        internal void RefreshFormula()
        {
            this.Formula = SLChartTool.GetChartReferenceFormula(this.WorksheetName, this.StartRowIndex, this.StartColumnIndex, this.EndRowIndex, this.EndColumnIndex);
        }

        internal SLStringReference Clone()
        {
            SLStringReference sr = new SLStringReference();
            sr.WorksheetName = this.WorksheetName;
            sr.StartRowIndex = this.StartRowIndex;
            sr.StartColumnIndex = this.StartColumnIndex;
            sr.EndRowIndex = this.EndRowIndex;
            sr.EndColumnIndex = this.EndColumnIndex;
            sr.Formula = this.Formula;
            sr.PointCount = this.PointCount;
            for (int i = 0; i < this.Points.Count; ++i)
            {
                sr.Points.Add(this.Points[i].Clone());
            }

            return sr;
        }
    }
}
