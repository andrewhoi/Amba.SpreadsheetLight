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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Amba.SpreadsheetLight
{
    internal class SLRowItem
    {
        internal List<int> MemberPropertyIndexes { get; set; }

        internal ItemValues ItemType { get; set; }
        internal uint RepeatedItemCount { get; set; }
        internal uint Index { get; set; }

        internal SLRowItem()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.ItemType = ItemValues.Data;
            this.RepeatedItemCount = 0;
            this.Index = 0;
        }

        internal void FromRowItem(RowItem ri)
        {
            this.SetAllNull();

            if (ri.ItemType != null) this.ItemType = ri.ItemType.Value;
            if (ri.RepeatedItemCount != null) this.RepeatedItemCount = ri.RepeatedItemCount.Value;
            if (ri.Index != null) this.Index = ri.Index.Value;

            MemberPropertyIndex mpi;
            using (OpenXmlReader oxr = OpenXmlReader.Create(ri))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(MemberPropertyIndex))
                    {
                        mpi = (MemberPropertyIndex)oxr.LoadCurrentElement();
                        if (mpi.Val != null) this.MemberPropertyIndexes.Add(mpi.Val.Value);
                        else this.MemberPropertyIndexes.Add(0);
                    }
                }
            }
        }

        internal RowItem ToRowItem()
        {
            RowItem ri = new RowItem();
            if (this.ItemType != ItemValues.Data) ri.ItemType = this.ItemType;
            if (this.RepeatedItemCount != 0) ri.RepeatedItemCount = this.RepeatedItemCount;
            if (this.Index != 0) ri.Index = this.Index;

            foreach (int i in this.MemberPropertyIndexes)
            {
                if (i != 0) ri.Append(new MemberPropertyIndex() { Val = i });
                else ri.Append(new MemberPropertyIndex());
            }

            return ri;
        }

        internal SLRowItem Clone()
        {
            SLRowItem ri = new SLRowItem();
            ri.ItemType = this.ItemType;
            ri.RepeatedItemCount = this.RepeatedItemCount;
            ri.Index = this.Index;

            ri.MemberPropertyIndexes = new List<int>();
            foreach (int i in this.MemberPropertyIndexes)
            {
                ri.MemberPropertyIndexes.Add(i);
            }

            return ri;
        }
    }
}
