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
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Amba.SpreadsheetLight
{
    internal class SLConditionalFormatValueObject
    {
        internal ConditionalFormatValueObjectValues Type { get; set; }
        internal string Val { get; set; }
        internal bool GreaterThanOrEqual { get; set; }

        internal SLConditionalFormatValueObject()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Type = ConditionalFormatValueObjectValues.Percentile;
            this.Val = string.Empty;
            this.GreaterThanOrEqual = true;
        }

        internal void FromConditionalFormatValueObject(ConditionalFormatValueObject cfvo)
        {
            this.SetAllNull();

            this.Type = cfvo.Type.Value;
            if (cfvo.Val != null) this.Val = cfvo.Val.Value;
            if (cfvo.GreaterThanOrEqual != null) this.GreaterThanOrEqual = cfvo.GreaterThanOrEqual.Value;
        }

        internal ConditionalFormatValueObject ToConditionalFormatValueObject()
        {
            ConditionalFormatValueObject cfvo = new ConditionalFormatValueObject();
            cfvo.Type = this.Type;

            if (this.Val.Length > 0)
            {
                if (this.Val.StartsWith("=")) cfvo.Val = this.Val.Substring(1);
                else cfvo.Val = this.Val;
            }
            
            if (!this.GreaterThanOrEqual) cfvo.GreaterThanOrEqual = false;

            return cfvo;
        }

        internal SLConditionalFormattingValueObject2010 ToSLConditionalFormattingValueObject2010()
        {
            SLConditionalFormattingValueObject2010 cfvo2010 = new SLConditionalFormattingValueObject2010();
            cfvo2010.Formula = this.Val;

            switch (this.Type)
            {
                case ConditionalFormatValueObjectValues.Formula:
                    cfvo2010.Type = X14.ConditionalFormattingValueObjectTypeValues.Formula;
                    break;
                case ConditionalFormatValueObjectValues.Max:
                    cfvo2010.Type = X14.ConditionalFormattingValueObjectTypeValues.Max;
                    break;
                case ConditionalFormatValueObjectValues.Min:
                    cfvo2010.Type = X14.ConditionalFormattingValueObjectTypeValues.Min;
                    break;
                case ConditionalFormatValueObjectValues.Number:
                    cfvo2010.Type = X14.ConditionalFormattingValueObjectTypeValues.Numeric;
                    break;
                case ConditionalFormatValueObjectValues.Percent:
                    cfvo2010.Type = X14.ConditionalFormattingValueObjectTypeValues.Percent;
                    break;
                case ConditionalFormatValueObjectValues.Percentile:
                    cfvo2010.Type = X14.ConditionalFormattingValueObjectTypeValues.Percentile;
                    break;
            }

            cfvo2010.GreaterThanOrEqual = this.GreaterThanOrEqual;

            return cfvo2010;
        }

        internal SLConditionalFormatValueObject Clone()
        {
            SLConditionalFormatValueObject cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = this.Type;
            cfvo.Val = this.Val;
            cfvo.GreaterThanOrEqual = this.GreaterThanOrEqual;

            return cfvo;
        }
    }
}
