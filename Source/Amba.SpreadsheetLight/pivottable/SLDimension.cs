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
    internal class SLDimension
    {
        internal bool Measure { get; set; }
        internal string Name { get; set; }
        internal string UniqueName { get; set; }
        internal string Caption { get; set; }

        internal SLDimension()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Measure = false;
            this.Name = "";
            this.UniqueName = "";
            this.Caption = "";
        }

        internal void FromDimension(Dimension d)
        {
            this.SetAllNull();

            if (d.Measure != null) this.Measure = d.Measure.Value;
            if (d.Name != null) this.Name = d.Name.Value;
            if (d.UniqueName != null) this.UniqueName = d.UniqueName.Value;
            if (d.Caption != null) this.Caption = d.Caption.Value;
        }

        internal Dimension ToDimension()
        {
            Dimension d = new Dimension();
            if (this.Measure != false) d.Measure = this.Measure;
            d.Name = this.Name;
            d.UniqueName = this.UniqueName;
            d.Caption = this.Caption;

            return d;
        }

        internal SLDimension Clone()
        {
            SLDimension d = new SLDimension();
            d.Measure = this.Measure;
            d.Name = this.Name;
            d.UniqueName = this.UniqueName;
            d.Caption = this.Caption;

            return d;
        }
    }
}
