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
    internal class SLTotalsRowFormula
    {
        internal bool Array { get; set; }
        internal string Text { get; set; }

        internal SLTotalsRowFormula()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Array = false;
            this.Text = string.Empty;
        }

        internal void FromTotalsRowFormula(TotalsRowFormula trf)
        {
            this.SetAllNull();

            if (trf.Array != null && trf.Array.Value) this.Array = true;
            this.Text = trf.Text;
        }

        internal TotalsRowFormula ToTotalsRowFormula()
        {
            TotalsRowFormula trf = new TotalsRowFormula();
            if (this.Array) trf.Array = this.Array;

            if (SLTool.ToPreserveSpace(this.Text))
            {
                trf.Space = SpaceProcessingModeValues.Preserve;
            }
            trf.Text = this.Text;

            return trf;
        }

        internal SLTotalsRowFormula Clone()
        {
            SLTotalsRowFormula trf = new SLTotalsRowFormula();
            trf.Array = this.Array;
            trf.Text = this.Text;

            return trf;
        }
    }
}
