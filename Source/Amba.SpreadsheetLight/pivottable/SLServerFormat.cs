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
    internal class SLServerFormat
    {
        internal string Culture { get; set; }
        internal string Format { get; set; }

        internal SLServerFormat()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Culture = "";
            this.Format = "";
        }

        internal void FromServerFormat(ServerFormat sf)
        {
            this.SetAllNull();

            if (sf.Culture != null) this.Culture = sf.Culture.Value;
            if (sf.Format != null) this.Format = sf.Format.Value;
        }

        internal ServerFormat ToServerFormat()
        {
            ServerFormat sf = new ServerFormat();
            if (this.Culture != null && this.Culture.Length > 0) sf.Culture = this.Culture;
            if (this.Format != null && this.Format.Length > 0) sf.Format = this.Format;

            return sf;
        }

        internal SLServerFormat Clone()
        {
            SLServerFormat sf = new SLServerFormat();
            sf.Culture = this.Culture;
            sf.Format = this.Format;

            return sf;
        }
    }
}
