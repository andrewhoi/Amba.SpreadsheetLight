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
    internal class SLPageField
    {
        internal int Field { get; set; }
        internal uint? Item { get; set; }
        internal int Hierarchy { get; set; }
        internal string Name { get; set; }
        internal string Caption { get; set; }

        internal SLPageField()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Field = 0;
            this.Item = null;
            this.Hierarchy = 0;
            this.Name = "";
            this.Caption = "";
        }

        internal void FromPageField(PageField pf)
        {
            this.SetAllNull();

            if (pf.Field != null) this.Field = pf.Field.Value;
            if (pf.Item != null) this.Item = pf.Item.Value;
            if (pf.Hierarchy != null) this.Hierarchy = pf.Hierarchy.Value;
            if (pf.Name != null) this.Name = pf.Name.Value;
            if (pf.Caption != null) this.Caption = pf.Caption.Value;
        }

        internal PageField ToPageField()
        {
            PageField pf = new PageField();
            pf.Field = this.Field;
            if (this.Item != null) pf.Item = this.Item.Value;
            pf.Hierarchy = this.Hierarchy;
            if (this.Name != null && this.Name.Length > 0) pf.Name = this.Name;
            if (this.Caption != null && this.Caption.Length > 0) pf.Caption = this.Caption;

            return pf;
        }

        internal SLPageField Clone()
        {
            SLPageField pf = new SLPageField();
            pf.Field = this.Field;
            pf.Item = this.Item;
            pf.Hierarchy = this.Hierarchy;
            pf.Name = this.Name;
            pf.Caption = this.Caption;

            return pf;
        }
    }
}
