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
    internal class SLColorFilter
    {
        internal uint? FormatId { get; set; }
        internal bool? CellColor { get; set; }

        internal SLColorFilter()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.FormatId = null;
            this.CellColor = null;
        }

        internal void FromColorFilter(ColorFilter cf)
        {
            this.SetAllNull();

            if (cf.FormatId != null) this.FormatId = cf.FormatId.Value;
            if (cf.CellColor != null && !cf.CellColor.Value) this.CellColor = cf.CellColor.Value;
        }

        internal ColorFilter ToColorFilter()
        {
            ColorFilter cf = new ColorFilter();
            if (this.FormatId != null) cf.FormatId = this.FormatId.Value;
            if (this.CellColor != null && !this.CellColor.Value) cf.CellColor = this.CellColor.Value;

            return cf;
        }

        internal SLColorFilter Clone()
        {
            SLColorFilter cf = new SLColorFilter();
            cf.FormatId = this.FormatId;
            cf.CellColor = this.CellColor;

            return cf;
        }
    }
}
