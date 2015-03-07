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
using System.Text;
using DocumentFormat.OpenXml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Excel = DocumentFormat.OpenXml.Office.Excel;

namespace Amba.SpreadsheetLight
{
    internal class SLSparkline
    {
        internal string WorksheetName;
        internal int StartRowIndex;
        internal int StartColumnIndex;
        internal int EndRowIndex;
        internal int EndColumnIndex;
        internal int LocationRowIndex;
        internal int LocationColumnIndex;

        internal SLSparkline()
        {
            this.WorksheetName = string.Empty;
            this.StartRowIndex = 1;
            this.StartColumnIndex = 1;
            this.EndRowIndex = 1;
            this.EndColumnIndex = 1;
            this.LocationRowIndex = 1;
            this.LocationColumnIndex = 1;
        }

        internal X14.Sparkline ToSparkline()
        {
            X14.Sparkline spk = new X14.Sparkline();

            if (this.StartRowIndex == this.EndRowIndex && this.StartColumnIndex == this.EndColumnIndex)
            {
                spk.Formula = new Excel.Formula();
                spk.Formula.Text = SLTool.ToCellReference(this.WorksheetName, this.StartRowIndex, this.StartColumnIndex);
            }
            else
            {
                spk.Formula = new Excel.Formula();
                spk.Formula.Text = SLTool.ToCellRange(this.WorksheetName, this.StartRowIndex, this.StartColumnIndex, this.EndRowIndex, this.EndColumnIndex);
            }

            spk.ReferenceSequence = new Excel.ReferenceSequence();
            spk.ReferenceSequence.Text = SLTool.ToCellReference(this.LocationRowIndex, this.LocationColumnIndex);

            return spk;
        }

        internal SLSparkline Clone()
        {
            SLSparkline spk = new SLSparkline();
            spk.WorksheetName = this.WorksheetName;
            spk.StartRowIndex = this.StartRowIndex;
            spk.StartColumnIndex = this.StartColumnIndex;
            spk.EndRowIndex = this.EndRowIndex;
            spk.EndColumnIndex = this.EndColumnIndex;
            spk.LocationRowIndex = this.LocationRowIndex;
            spk.LocationColumnIndex = this.LocationColumnIndex;

            return spk;
        }
    }
}
