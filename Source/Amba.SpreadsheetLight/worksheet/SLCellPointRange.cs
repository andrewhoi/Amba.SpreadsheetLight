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
    internal struct SLCellPointRange
    {
        internal int StartRowIndex;
        internal int StartColumnIndex;
        internal int EndRowIndex;
        internal int EndColumnIndex;

        internal SLCellPointRange(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            this.StartRowIndex = StartRowIndex;
            this.StartColumnIndex = StartColumnIndex;
            this.EndRowIndex = EndRowIndex;
            this.EndColumnIndex = EndColumnIndex;
        }
    }

    internal class SLCellPointRangeComparer : IComparer<SLCellPointRange>
    {
        public int Compare(SLCellPointRange pt1, SLCellPointRange pt2)
        {
            if (pt1.StartRowIndex < pt2.StartRowIndex)
            {
                return -1;
            }
            else if (pt1.StartRowIndex > pt2.StartRowIndex)
            {
                return 1;
            }
            else
            {
                if (pt1.StartColumnIndex < pt2.StartColumnIndex)
                {
                    return -1;
                }
                else if (pt1.StartColumnIndex > pt2.StartColumnIndex)
                {
                    return 1;
                }
                else
                {
                    if (pt1.EndRowIndex < pt2.EndRowIndex)
                    {
                        return -1;
                    }
                    else if (pt1.EndRowIndex > pt2.EndRowIndex)
                    {
                        return 1;
                    }
                    else
                    {
                        return pt1.EndColumnIndex.CompareTo(pt2.EndColumnIndex);
                    }
                }
            }
        }
    }
}
