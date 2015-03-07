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
    internal class SLDynamicFilter
    {
        internal DynamicFilterValues Type { get; set; }
        internal double? Val { get; set; }
        internal double? MaxVal { get; set; }

        internal SLDynamicFilter()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Type = DynamicFilterValues.Null;
            this.Val = null;
            this.MaxVal = null;
        }

        internal void FromDynamicFilter(DynamicFilter df)
        {
            this.SetAllNull();

            this.Type = df.Type.Value;
            if (df.Val != null) this.Val = df.Val.Value;
            if (df.MaxVal != null) this.MaxVal = df.MaxVal.Value;
        }

        internal DynamicFilter ToDynamicFilter()
        {
            DynamicFilter df = new DynamicFilter();
            df.Type = this.Type;
            if (this.Val != null) df.Val = this.Val.Value;
            if (this.MaxVal != null) df.MaxVal = this.MaxVal.Value;

            return df;
        }

        internal SLDynamicFilter Clone()
        {
            SLDynamicFilter df = new SLDynamicFilter();
            df.Type = this.Type;
            df.Val = this.Val;
            df.MaxVal = this.MaxVal;

            return df;
        }
    }
}
