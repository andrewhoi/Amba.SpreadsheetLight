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
using System.Collections.Generic;

namespace Amba.SpreadsheetLight
{
    internal struct SLSortItem
    {
        internal double Number;
        internal string Text;
        internal int Index;

        internal SLSortItem(double SortNumber, string SortText, int SortIndex)
        {
            Number = SortNumber;
            Text = SortText;
            Index = SortIndex;
        }
    }

    internal class SLSortItemNumberComparer : IComparer<SLSortItem>
    {
        public int Compare(SLSortItem si1, SLSortItem si2)
        {
            return si1.Number.CompareTo(si2.Number);
        }
    }

    internal class SLSortItemTextComparer : IComparer<SLSortItem>
    {
        public int Compare(SLSortItem si1, SLSortItem si2)
        {
            return si1.Text.CompareTo(si2.Text);
        }
    }
}
