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
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Amba.SpreadsheetLight
{
    internal class SLConditionalFormattingIcon2010
    {
        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.conditionalformattingicon.aspx

        internal X14.IconSetTypeValues IconSet { get; set; }
        internal uint IconId { get; set; }

        internal SLConditionalFormattingIcon2010()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.IconSet = X14.IconSetTypeValues.ThreeTrafficLights1;
            this.IconId = 0;
        }

        internal void FromConditionalFormattingIcon(X14.ConditionalFormattingIcon cfi)
        {
            this.SetAllNull();

            if (cfi.IconSet != null) this.IconSet = cfi.IconSet.Value;
            if (cfi.IconId != null) this.IconId = cfi.IconId.Value;
        }

        internal X14.ConditionalFormattingIcon ToConditionalFormattingIcon()
        {
            X14.ConditionalFormattingIcon cfi = new X14.ConditionalFormattingIcon();

            cfi.IconSet = this.IconSet;
            cfi.IconId = this.IconId;

            return cfi;
        }

        internal SLConditionalFormattingIcon2010 Clone()
        {
            SLConditionalFormattingIcon2010 cfi = new SLConditionalFormattingIcon2010();
            cfi.IconSet = this.IconSet;
            cfi.IconId = this.IconId;

            return cfi;
        }
    }
}
