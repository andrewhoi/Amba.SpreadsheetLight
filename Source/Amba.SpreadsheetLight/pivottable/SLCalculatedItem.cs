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
    internal class SLCalculatedItem
    {
        internal SLPivotArea PivotArea { get; set; }

        internal uint? Field { get; set; }
        internal string Formula { get; set; }

        internal SLCalculatedItem()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.PivotArea = new SLPivotArea();
            this.Field = null;
            this.Formula = "";
        }

        internal void FromCalculatedItem(CalculatedItem ci)
        {
            this.SetAllNull();

            if (ci.Field != null) this.Field = ci.Field.Value;
            if (ci.Formula != null) this.Formula = ci.Formula.Value;

            if (ci.PivotArea != null) this.PivotArea.FromPivotArea(ci.PivotArea);
        }

        internal CalculatedItem ToCalculatedItem()
        {
            CalculatedItem ci = new CalculatedItem();
            if (this.Field != null) ci.Field = this.Field.Value;
            if (this.Formula != null && this.Formula.Length > 0) ci.Formula = this.Formula;

            ci.PivotArea = this.PivotArea.ToPivotArea();

            return ci;
        }

        internal SLCalculatedItem Clone()
        {
            SLCalculatedItem ci = new SLCalculatedItem();
            ci.Field = this.Field;
            ci.Formula = this.Formula;
            ci.PivotArea = this.PivotArea.Clone();

            return ci;
        }
    }
}
