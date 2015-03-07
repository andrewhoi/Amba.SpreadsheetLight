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
    internal class SLRangeProperties
    {
        internal bool AutoStart { get; set; }
        internal bool AutoEnd { get; set; }
        internal GroupByValues GroupBy { get; set; }
        internal double? StartNumber { get; set; }
        internal double? EndNum { get; set; }
        internal DateTime? StartDate { get; set; }
        internal DateTime? EndDate { get; set; }
        internal double GroupInterval { get; set; }

        internal SLRangeProperties()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.AutoStart = true;
            this.AutoEnd = true;
            this.GroupBy = GroupByValues.Range;
            this.StartNumber = null;
            this.EndNum = null;
            this.StartDate = null;
            this.EndDate = null;
            this.GroupInterval = 1;
        }

        internal void FromRangeProperties(RangeProperties rp)
        {
            this.SetAllNull();

            if (rp.AutoStart != null) this.AutoStart = rp.AutoStart.Value;
            if (rp.AutoEnd != null) this.AutoEnd = rp.AutoEnd.Value;
            if (rp.GroupBy != null) this.GroupBy = rp.GroupBy.Value;
            if (rp.StartNumber != null) this.StartNumber = rp.StartNumber.Value;
            if (rp.EndNum != null) this.EndNum = rp.EndNum.Value;
            if (rp.StartDate != null) this.StartDate = rp.StartDate.Value;
            if (rp.EndDate != null) this.EndDate = rp.EndDate.Value;
            if (rp.GroupInterval != null) this.GroupInterval = rp.GroupInterval.Value;
        }

        internal RangeProperties ToRangeProperties()
        {
            RangeProperties rp = new RangeProperties();
            if (this.AutoStart != true) rp.AutoStart = this.AutoStart;
            if (this.AutoEnd != true) rp.AutoEnd = this.AutoEnd;
            if (this.GroupBy != GroupByValues.Range) rp.GroupBy = this.GroupBy;
            if (this.StartNumber != null) rp.StartNumber = this.StartNumber.Value;
            if (this.EndNum != null) rp.EndNum = this.EndNum.Value;
            if (this.StartDate != null) rp.StartDate = this.StartDate.Value;
            if (this.EndDate != null) rp.EndDate = this.EndDate.Value;
            if (this.GroupInterval != 1) rp.GroupInterval = this.GroupInterval;

            return rp;
        }

        internal SLRangeProperties Clone()
        {
            SLRangeProperties rp = new SLRangeProperties();
            rp.AutoStart = this.AutoStart;
            rp.AutoEnd = this.AutoEnd;
            rp.GroupBy = this.GroupBy;
            rp.StartNumber = this.StartNumber;
            rp.EndNum = this.EndNum;
            rp.StartDate = this.StartDate;
            rp.EndDate = this.EndDate;
            rp.GroupInterval = this.GroupInterval;

            return rp;
        }
    }
}
