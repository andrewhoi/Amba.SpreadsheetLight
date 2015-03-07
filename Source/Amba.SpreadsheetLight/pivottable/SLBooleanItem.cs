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
    internal class SLBooleanItem
    {
        internal List<int> MemberPropertyIndexes { get; set; }

        internal bool Val { get; set; }
        internal bool? Unused { get; set; }
        internal bool? Calculated { get; set; }
        internal string Caption { get; set; }
        internal uint? PropertyCount { get; set; }

        internal SLBooleanItem()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.MemberPropertyIndexes = new List<int>();

            this.Val = true;
            this.Unused = null;
            this.Calculated = null;
            this.Caption = "";
            this.PropertyCount = null;
        }

        internal void FromBooleanItem(BooleanItem bi)
        {
            this.SetAllNull();

            if (bi.Val != null) this.Val = bi.Val.Value;
            if (bi.Unused != null) this.Unused = bi.Unused.Value;
            if (bi.Calculated != null) this.Calculated = bi.Calculated.Value;
            if (bi.Caption != null) this.Caption = bi.Caption.Value;
            if (bi.PropertyCount != null) this.PropertyCount = bi.PropertyCount.Value;

            MemberPropertyIndex mpi;
            using (OpenXmlReader oxr = OpenXmlReader.Create(bi))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(MemberPropertyIndex))
                    {
                        // 0 is the default value.
                        mpi = (MemberPropertyIndex)oxr.LoadCurrentElement();
                        if (mpi.Val != null) this.MemberPropertyIndexes.Add(mpi.Val.Value);
                        else this.MemberPropertyIndexes.Add(0);
                    }
                }
            }
        }

        internal BooleanItem ToBooleanItem()
        {
            BooleanItem bi = new BooleanItem();
            bi.Val = this.Val;
            if (this.Unused != null) bi.Unused = this.Unused.Value;
            if (this.Calculated != null) bi.Calculated = this.Calculated.Value;
            if (this.Caption != null && this.Caption.Length > 0) bi.Caption = this.Caption;
            if (this.PropertyCount != null) bi.PropertyCount = this.PropertyCount.Value;

            foreach (int i in this.MemberPropertyIndexes)
            {
                if (i != 0) bi.Append(new MemberPropertyIndex() { Val = i });
                else bi.Append(new MemberPropertyIndex());
            }

            return bi;
        }

        internal SLBooleanItem Clone()
        {
            SLBooleanItem bi = new SLBooleanItem();
            bi.Val = this.Val;
            bi.Unused = this.Unused;
            bi.Calculated = this.Calculated;
            bi.Caption = this.Caption;
            bi.PropertyCount = this.PropertyCount;

            bi.MemberPropertyIndexes = new List<int>();
            foreach (int i in this.MemberPropertyIndexes)
            {
                bi.MemberPropertyIndexes.Add(i);
            }

            return bi;
        }
    }
}
