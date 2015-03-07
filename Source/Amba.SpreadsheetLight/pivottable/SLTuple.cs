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
using DocumentFormat.OpenXml.Spreadsheet;

// Apparently .NET Framework 4 has a System.Tuple, which clashes
// with DocumentFormat.OpenXml.Spreadsheet.Tuple.
// Good thing we're on 3.5...

namespace Amba.SpreadsheetLight
{
    internal class SLTuple
    {
        internal uint? Field { get; set; }
        internal uint? Hierarchy { get; set; }
        internal uint Item { get; set; }

        internal SLTuple()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Field = null;
            this.Hierarchy = null;
            this.Item = 0;
        }

        internal void FromTuple(Tuple t)
        {
            this.SetAllNull();

            if (t.Field != null) this.Field = t.Field.Value;
            if (t.Hierarchy != null) this.Hierarchy = t.Hierarchy.Value;
            if (t.Item != null) this.Item = t.Item.Value;
        }

        internal Tuple ToTuple()
        {
            Tuple t = new Tuple();
            if (this.Field != null) t.Field = this.Field.Value;
            if (this.Hierarchy != null) t.Hierarchy = this.Hierarchy.Value;
            t.Item = this.Item;

            return t;
        }

        internal SLTuple Clone()
        {
            SLTuple t = new SLTuple();
            t.Field = this.Field;
            t.Hierarchy = this.Hierarchy;
            t.Item = this.Item;

            return t;
        }
    }
}
