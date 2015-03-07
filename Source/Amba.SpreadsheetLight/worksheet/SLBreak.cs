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
    internal class SLBreak
    {
        internal uint Id { get; set; }
        internal uint Min { get; set; }
        internal uint Max { get; set; }
        internal bool ManualPageBreak { get; set; }
        internal bool PivotTablePageBreak { get; set; }

        internal SLBreak()
        {
            this.SetAllNull();
        }

        internal void SetAllNull()
        {
            this.Id = 0;
            this.Min = 0;
            this.Max = 0;
            this.ManualPageBreak = false;
            this.PivotTablePageBreak = false;
        }

        internal void FromBreak(Break b)
        {
            this.SetAllNull();
            if (b.Id != null) this.Id = b.Id;
            if (b.Min != null) this.Min = b.Min;
            if (b.Max != null) this.Max = b.Max;
            if (b.ManualPageBreak != null) this.ManualPageBreak = b.ManualPageBreak;
            if (b.PivotTablePageBreak != null) this.PivotTablePageBreak = b.PivotTablePageBreak;
        }

        internal Break ToBreak()
        {
            Break b = new Break();
            if (this.Id != 0) b.Id = this.Id;
            if (this.Min != 0) b.Min = this.Min;
            if (this.Max != 0) b.Max = this.Max;
            if (this.ManualPageBreak != false) b.ManualPageBreak = this.ManualPageBreak;
            if (this.PivotTablePageBreak != false) b.PivotTablePageBreak = this.PivotTablePageBreak;

            return b;
        }

        internal SLBreak Clone()
        {
            SLBreak b = new SLBreak();
            b.Id = this.Id;
            b.Min = this.Min;
            b.Max = this.Max;
            b.ManualPageBreak = this.ManualPageBreak;
            b.PivotTablePageBreak = this.PivotTablePageBreak;

            return b;
        }
    }
}
