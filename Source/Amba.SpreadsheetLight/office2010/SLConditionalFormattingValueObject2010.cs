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
using Excel = DocumentFormat.OpenXml.Office.Excel;

namespace Amba.SpreadsheetLight
{
    internal class SLConditionalFormattingValueObject2010
    {
        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.conditionalformattingvalueobject.aspx

        internal string Formula { get; set; }
        internal X14.ConditionalFormattingValueObjectTypeValues Type { get; set; }
        internal bool GreaterThanOrEqual { get; set; }
        
        internal SLConditionalFormattingValueObject2010()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Formula = string.Empty;
            this.Type = X14.ConditionalFormattingValueObjectTypeValues.Percentile;
            this.GreaterThanOrEqual = true;
        }

        internal void FromConditionalFormattingValueObject(X14.ConditionalFormattingValueObject cfvo)
        {
            this.SetAllNull();

            if (cfvo.Formula != null) this.Formula = cfvo.Formula.Text;
            this.Type = cfvo.Type.Value;
            if (cfvo.GreaterThanOrEqual != null) this.GreaterThanOrEqual = cfvo.GreaterThanOrEqual.Value;
        }

        internal X14.ConditionalFormattingValueObject ToConditionalFormattingValueObject()
        {
            X14.ConditionalFormattingValueObject cfvo = new X14.ConditionalFormattingValueObject();

            if (this.Formula.Length > 0)
            {
                if (this.Formula.StartsWith("="))
                {
                    cfvo.Formula = new Excel.Formula(this.Formula.Substring(1));
                }
                else
                {
                    cfvo.Formula = new Excel.Formula(this.Formula);
                }
            }
            cfvo.Type = this.Type;
            if (!this.GreaterThanOrEqual) cfvo.GreaterThanOrEqual = this.GreaterThanOrEqual;

            return cfvo;
        }

        internal SLConditionalFormattingValueObject2010 Clone()
        {
            SLConditionalFormattingValueObject2010 cfvo = new SLConditionalFormattingValueObject2010();
            cfvo.Formula = this.Formula;
            cfvo.Type = this.Type;
            cfvo.GreaterThanOrEqual = this.GreaterThanOrEqual;

            return cfvo;
        }
    }
}
