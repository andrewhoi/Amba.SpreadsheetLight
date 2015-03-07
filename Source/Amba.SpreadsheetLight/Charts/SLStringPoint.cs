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
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Amba.SpreadsheetLight.Charts
{
    internal class SLStringPoint
    {
        internal string NumericValue { get; set; }
        internal uint Index { get; set; }

        internal SLStringPoint()
        {
            this.NumericValue = string.Empty;
            this.Index = 0;
        }

        internal C.StringPoint ToStringPoint()
        {
            C.StringPoint sp = new C.StringPoint();
            sp.Index = this.Index;
            sp.NumericValue = new C.NumericValue(this.NumericValue);

            return sp;
        }

        internal SLStringPoint Clone()
        {
            SLStringPoint sp = new SLStringPoint();
            sp.NumericValue = this.NumericValue;
            sp.Index = this.Index;

            return sp;
        }
    }
}
