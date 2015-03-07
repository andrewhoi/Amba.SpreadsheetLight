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
    internal class SLDateGroupItem
    {
        internal ushort Year { get; set; }
        internal ushort? Month { get; set; }
        internal ushort? Day { get; set; }
        internal ushort? Hour { get; set; }
        internal ushort? Minute { get; set; }
        internal ushort? Second { get; set; }
        internal DateTimeGroupingValues DateTimeGrouping { get; set; }

        internal SLDateGroupItem()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Year = (ushort)DateTime.Now.Year;
            this.Month = null;
            this.Day = null;
            this.Hour = null;
            this.Minute = null;
            this.Second = null;
            this.DateTimeGrouping = DateTimeGroupingValues.Year;
        }

        internal void FromDateGroupItem(DateGroupItem dgi)
        {
            this.SetAllNull();

            this.Year = dgi.Year.Value;
            if (dgi.Month != null) this.Month = dgi.Month.Value;
            if (dgi.Day != null) this.Day = dgi.Day.Value;
            if (dgi.Hour != null) this.Hour = dgi.Hour.Value;
            if (dgi.Minute != null) this.Minute = dgi.Minute.Value;
            if (dgi.Second != null) this.Second = dgi.Second.Value;
            this.DateTimeGrouping = dgi.DateTimeGrouping.Value;
        }

        internal DateGroupItem ToDateGroupItem()
        {
            DateGroupItem dgi = new DateGroupItem();
            dgi.Year = this.Year;
            if (this.Month != null) dgi.Month = this.Month.Value;
            if (this.Day != null) dgi.Day = this.Day.Value;
            if (this.Hour != null) dgi.Hour = this.Hour.Value;
            if (this.Minute != null) dgi.Minute = this.Minute.Value;
            if (this.Second != null) dgi.Second = this.Second.Value;
            dgi.DateTimeGrouping = this.DateTimeGrouping;

            return dgi;
        }

        internal SLDateGroupItem Clone()
        {
            SLDateGroupItem dgi = new SLDateGroupItem();
            dgi.Year = this.Year;
            dgi.Month = this.Month;
            dgi.Day = this.Day;
            dgi.Hour = this.Hour;
            dgi.Minute = this.Minute;
            dgi.Second = this.Second;
            dgi.DateTimeGrouping = this.DateTimeGrouping;

            return dgi;
        }
    }
}
