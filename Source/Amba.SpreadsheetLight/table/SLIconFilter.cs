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
    internal class SLIconFilter
    {
        internal IconSetValues IconSet { get; set; }
        internal uint? IconId { get; set; }

        internal SLIconFilter()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.IconSet = IconSetValues.ThreeArrows;
            this.IconId = null;
        }

        internal void FromIconFilter(IconFilter icf)
        {
            this.SetAllNull();

            this.IconSet = icf.IconSet.Value;
            if (icf.IconId != null) this.IconId = icf.IconId.Value;
        }

        internal IconFilter ToIconFilter()
        {
            IconFilter icf = new IconFilter();
            icf.IconSet = this.IconSet;
            if (this.IconId != null) icf.IconId = this.IconId.Value;

            return icf;
        }

        internal SLIconFilter Clone()
        {
            SLIconFilter icf = new SLIconFilter();
            icf.IconSet = this.IconSet;
            icf.IconId = this.IconId;

            return icf;
        }
    }
}
