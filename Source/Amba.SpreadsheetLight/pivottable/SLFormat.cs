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
    internal class SLFormat
    {
        internal SLPivotArea PivotArea { get; set; }
        internal FormatActionValues Action { get; set; }
        internal uint? FormatId { get; set; }

        internal SLFormat()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.PivotArea = new SLPivotArea();
            this.Action = FormatActionValues.Formatting;
            this.FormatId = null;
        }

        internal void FromFormat(Format f)
        {
            this.SetAllNull();

            if (f.PivotArea != null) this.PivotArea.FromPivotArea(f.PivotArea);

            if (f.Action != null) this.Action = f.Action.Value;
            if (f.FormatId != null) this.FormatId = f.FormatId.Value;
        }

        internal Format ToFormat()
        {
            Format f = new Format();
            f.PivotArea = this.PivotArea.ToPivotArea();

            if (this.Action != FormatActionValues.Formatting) f.Action = this.Action;
            if (this.FormatId != null) f.FormatId = this.FormatId.Value;

            return f;
        }

        internal SLFormat Clone()
        {
            SLFormat f = new SLFormat();
            f.PivotArea = this.PivotArea.Clone();

            f.Action = this.Action;
            f.FormatId = this.FormatId;

            return f;
        }
    }
}
