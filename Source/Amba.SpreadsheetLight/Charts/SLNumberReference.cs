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
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Amba.SpreadsheetLight.Charts
{
    internal class SLNumberReference
    {
        internal string WorksheetName { get; set; }
        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        internal string Formula { get; set; }
        internal SLNumberingCache NumberingCache { get; set; }

        internal SLNumberReference()
        {
            this.WorksheetName = string.Empty;
            this.StartRowIndex = 1;
            this.StartColumnIndex = 1;
            this.EndRowIndex = 1;
            this.EndColumnIndex = 1;

            this.Formula = string.Empty;
            this.NumberingCache = new SLNumberingCache();
        }

        internal C.NumberReference ToNumberReference()
        {
            C.NumberReference nr = new C.NumberReference();
            nr.Formula = new C.Formula(this.Formula);
            nr.NumberingCache = this.NumberingCache.ToNumberingCache();

            return nr;
        }

        internal void RefreshFormula()
        {
            this.Formula = SLChartTool.GetChartReferenceFormula(this.WorksheetName, this.StartRowIndex, this.StartColumnIndex, this.EndRowIndex, this.EndColumnIndex);
        }

        internal SLNumberReference Clone()
        {
            SLNumberReference nr = new SLNumberReference();
            nr.WorksheetName = this.WorksheetName;
            nr.StartRowIndex = this.StartRowIndex;
            nr.StartColumnIndex = this.StartColumnIndex;
            nr.EndRowIndex = this.EndRowIndex;
            nr.EndColumnIndex = this.EndColumnIndex;
            nr.Formula = this.Formula;
            nr.NumberingCache = this.NumberingCache.Clone();

            return nr;
        }
    }
}
