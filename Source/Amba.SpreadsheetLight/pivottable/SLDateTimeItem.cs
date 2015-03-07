﻿// Copyright (c) 2011 Vincent Tan Wai Lip
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
    internal class SLDateTimeItem
    {
        internal List<int> MemberPropertyIndexes { get; set; }

        internal DateTime Val { get; set; }
        internal bool? Unused { get; set; }
        internal bool? Calculated { get; set; }
        internal string Caption { get; set; }
        internal uint? PropertyCount { get; set; }

        internal SLDateTimeItem()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.MemberPropertyIndexes = new List<int>();

            this.Val = new DateTime();
            this.Unused = null;
            this.Calculated = null;
            this.Caption = "";
            this.PropertyCount = null;
        }

        internal void FromDateTimeItem(DateTimeItem dti)
        {
            this.SetAllNull();

            if (dti.Val != null) this.Val = dti.Val.Value;
            if (dti.Unused != null) this.Unused = dti.Unused.Value;
            if (dti.Calculated != null) this.Calculated = dti.Calculated.Value;
            if (dti.Caption != null) this.Caption = dti.Caption.Value;
            if (dti.PropertyCount != null) this.PropertyCount = dti.PropertyCount.Value;

            MemberPropertyIndex mpi;
            using (OpenXmlReader oxr = OpenXmlReader.Create(dti))
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

        internal DateTimeItem ToDateTimeItem()
        {
            DateTimeItem dti = new DateTimeItem();
            dti.Val = this.Val;
            if (this.Unused != null) dti.Unused = this.Unused.Value;
            if (this.Calculated != null) dti.Calculated = this.Calculated.Value;
            if (this.Caption != null && this.Caption.Length > 0) dti.Caption = this.Caption;
            if (this.PropertyCount != null) dti.PropertyCount = this.PropertyCount.Value;

            foreach (int i in this.MemberPropertyIndexes)
            {
                if (i != 0) dti.Append(new MemberPropertyIndex() { Val = i });
                else dti.Append(new MemberPropertyIndex());
            }

            return dti;
        }

        internal SLDateTimeItem Clone()
        {
            SLDateTimeItem dti = new SLDateTimeItem();
            dti.Val = this.Val;
            dti.Unused = this.Unused;
            dti.Calculated = this.Calculated;
            dti.Caption = this.Caption;
            dti.PropertyCount = this.PropertyCount;

            dti.MemberPropertyIndexes = new List<int>();
            foreach (int i in this.MemberPropertyIndexes)
            {
                dti.MemberPropertyIndexes.Add(i);
            }

            return dti;
        }
    }
}
