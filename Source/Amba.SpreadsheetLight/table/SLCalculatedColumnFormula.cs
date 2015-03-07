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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Amba.SpreadsheetLight
{
    internal class SLCalculatedColumnFormula
    {
        internal bool Array { get; set; }
        internal string Text { get; set; }

        internal SLCalculatedColumnFormula()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Array = false;
            this.Text = string.Empty;
        }

        internal void FromCalculatedColumnFormula(CalculatedColumnFormula ccf)
        {
            this.SetAllNull();

            if (ccf.Array != null && ccf.Array.Value) this.Array = true;
            this.Text = ccf.Text;
        }

        internal CalculatedColumnFormula ToCalculatedColumnFormula()
        {
            CalculatedColumnFormula ccf = new CalculatedColumnFormula();
            if (this.Array) ccf.Array = this.Array;

            if (SLTool.ToPreserveSpace(this.Text))
            {
                ccf.Space = SpaceProcessingModeValues.Preserve;
            }
            ccf.Text = this.Text;

            return ccf;
        }

        internal SLCalculatedColumnFormula Clone()
        {
            SLCalculatedColumnFormula ccf = new SLCalculatedColumnFormula();
            ccf.Array = this.Array;
            ccf.Text = this.Text;

            return ccf;
        }
    }
}
