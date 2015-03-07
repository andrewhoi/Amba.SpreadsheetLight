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
    internal class SLTop10
    {
        internal bool? Top { get; set; }
        internal bool? Percent { get; set; }
        internal double Val { get; set; }
        internal double? FilterValue { get; set; }

        internal SLTop10()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Top = null;
            this.Percent = null;
            this.Val = 0.0;
            this.FilterValue = null;
        }

        internal void FromTop10(Top10 t)
        {
            this.SetAllNull();

            if (t.Top != null) this.Top = t.Top.Value;
            if (t.Percent != null) this.Percent = t.Percent.Value;
            this.Val = t.Val.Value;
            if (t.FilterValue != null) this.FilterValue = t.FilterValue.Value;
        }

        internal Top10 ToTop10()
        {
            Top10 t = new Top10();
            if (this.Top != null && !this.Top.Value) t.Top = this.Top.Value;
            if (this.Percent != null && this.Percent.Value) t.Percent = this.Percent.Value;
            t.Val = this.Val;
            if (this.FilterValue != null) t.FilterValue = this.FilterValue.Value;

            return t;
        }

        internal SLTop10 Clone()
        {
            SLTop10 t = new SLTop10();
            t.Top = this.Top;
            t.Percent = this.Percent;
            t.Val = this.Val;
            t.FilterValue = this.FilterValue;

            return t;
        }
    }
}
