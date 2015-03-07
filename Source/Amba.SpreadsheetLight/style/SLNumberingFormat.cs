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
    internal class SLNumberingFormat
    {
        private uint iNumberFormatId;
        internal uint NumberFormatId
        {
            get { return iNumberFormatId; }
            set { iNumberFormatId = value; }
        }

        private string sFormatCode;
        internal string FormatCode
        {
            get { return sFormatCode; }
            set { sFormatCode = value; }
        }

        internal SLNumberingFormat()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.iNumberFormatId = 0;
            this.sFormatCode = string.Empty;
        }

        internal void FromNumberingFormat(NumberingFormat nf)
        {
            this.SetAllNull();

            if (nf.NumberFormatId != null)
            {
                this.NumberFormatId = nf.NumberFormatId.Value;
            }
            else
            {
                this.NumberFormatId = 0;
            }

            if (nf.FormatCode != null)
            {
                this.FormatCode = nf.FormatCode.Value;
            }
            else
            {
                this.FormatCode = string.Empty;
            }
        }

        internal NumberingFormat ToNumberingFormat()
        {
            NumberingFormat nf = new NumberingFormat();
            nf.NumberFormatId = this.NumberFormatId;
            nf.FormatCode = this.FormatCode;

            return nf;
        }

        internal void FromHash(string Hash)
        {
            this.FormatCode = Hash;
        }

        internal string ToHash()
        {
            return this.FormatCode;
        }

        internal SLNumberingFormat Clone()
        {
            SLNumberingFormat nf = new SLNumberingFormat();
            nf.iNumberFormatId = this.iNumberFormatId;
            nf.sFormatCode = this.sFormatCode;

            return nf;
        }
    }
}
