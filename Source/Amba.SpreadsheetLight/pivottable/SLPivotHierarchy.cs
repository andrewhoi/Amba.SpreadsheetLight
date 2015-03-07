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
    internal class SLPivotHierarchy
    {
        internal List<SLMemberProperty> MemberProperties { get; set; }
        internal List<SLMembers> Members { get; set; }

        internal bool Outline { get; set; }
        internal bool MultipleItemSelectionAllowed { get; set; }
        internal bool SubtotalTop { get; set; }
        internal bool ShowInFieldList { get; set; }
        internal bool DragToRow { get; set; }
        internal bool DragToColumn { get; set; }
        internal bool DragToPage { get; set; }
        internal bool DragToData { get; set; }
        internal bool DragOff { get; set; }
        internal bool IncludeNewItemsInFilter { get; set; }
        internal string Caption { get; set; }

        internal SLPivotHierarchy()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.MemberProperties = new List<SLMemberProperty>();
            this.Members = new List<SLMembers>();

            this.Outline = false;
            this.MultipleItemSelectionAllowed = false;
            this.SubtotalTop = false;
            this.ShowInFieldList = true;
            this.DragToRow = true;
            this.DragToColumn = true;
            this.DragToPage = true;
            this.DragToData = false;
            this.DragOff = true;
            this.IncludeNewItemsInFilter = false;
            this.Caption = "";
        }

        internal void FromPivotHierarchy(PivotHierarchy ph)
        {
            this.SetAllNull();

            if (ph.Outline != null) this.Outline = ph.Outline.Value;
            if (ph.MultipleItemSelectionAllowed != null) this.Outline = ph.MultipleItemSelectionAllowed.Value;
            if (ph.SubtotalTop != null) this.SubtotalTop = ph.SubtotalTop.Value;
            if (ph.ShowInFieldList != null) this.ShowInFieldList = ph.ShowInFieldList.Value;
            if (ph.DragToRow != null) this.DragToRow = ph.DragToRow.Value;
            if (ph.DragToColumn != null) this.DragToColumn = ph.DragToColumn.Value;
            if (ph.DragToPage != null) this.DragToPage = ph.DragToPage.Value;
            if (ph.DragToData != null) this.DragToData = ph.DragToData.Value;
            if (ph.DragOff != null) this.DragOff = ph.DragOff.Value;
            if (ph.IncludeNewItemsInFilter != null) this.IncludeNewItemsInFilter = ph.IncludeNewItemsInFilter.Value;
            if (ph.Caption != null) this.Caption = ph.Caption.Value;

            SLMemberProperty mp;
            SLMembers mems;
            using (OpenXmlReader oxr = OpenXmlReader.Create(ph))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(MemberProperty))
                    {
                        mp = new SLMemberProperty();
                        mp.FromMemberProperty((MemberProperty)oxr.LoadCurrentElement());
                        this.MemberProperties.Add(mp);
                    }
                    else if (oxr.ElementType == typeof(Members))
                    {
                        mems = new SLMembers();
                        mems.FromMembers((Members)oxr.LoadCurrentElement());
                        this.Members.Add(mems);
                    }
                }
            }
        }

        internal PivotHierarchy ToPivotHierarchy()
        {
            PivotHierarchy ph = new PivotHierarchy();

            if (this.Outline != false) ph.Outline = this.Outline;
            if (this.MultipleItemSelectionAllowed != false) ph.MultipleItemSelectionAllowed = this.MultipleItemSelectionAllowed;
            if (this.SubtotalTop != false) ph.SubtotalTop = this.SubtotalTop;
            if (this.ShowInFieldList != true) ph.ShowInFieldList = this.ShowInFieldList;
            if (this.DragToRow != true) ph.DragToRow = this.DragToRow;
            if (this.DragToColumn != true) ph.DragToColumn = this.DragToColumn;
            if (this.DragToPage != true) ph.DragToPage = this.DragToPage;
            if (this.DragToData != false) ph.DragToData = this.DragToData;
            if (this.DragOff != true) ph.DragOff = this.DragOff;
            if (this.IncludeNewItemsInFilter != false) ph.IncludeNewItemsInFilter = this.IncludeNewItemsInFilter;
            if (this.Caption != null && this.Caption.Length > 0) ph.Caption = this.Caption;

            if (this.MemberProperties.Count > 0)
            {
                ph.MemberProperties = new MemberProperties() { Count = (uint)this.MemberProperties.Count };
                foreach (SLMemberProperty mp in this.MemberProperties)
                {
                    ph.MemberProperties.Append(mp.ToMemberProperty());
                }
            }

            foreach (SLMembers mems in this.Members)
            {
                ph.Append(mems.ToMembers());
            }

            return ph;
        }

        internal SLPivotHierarchy Clone()
        {
            SLPivotHierarchy ph = new SLPivotHierarchy();
            ph.Outline = this.Outline;
            ph.MultipleItemSelectionAllowed = this.MultipleItemSelectionAllowed;
            ph.SubtotalTop = this.SubtotalTop;
            ph.ShowInFieldList = this.ShowInFieldList;
            ph.DragToRow = this.DragToRow;
            ph.DragToColumn = this.DragToColumn;
            ph.DragToPage = this.DragToPage;
            ph.DragToData = this.DragToData;
            ph.DragOff = this.DragOff;
            ph.IncludeNewItemsInFilter = this.IncludeNewItemsInFilter;
            ph.Caption = this.Caption;

            ph.MemberProperties = new List<SLMemberProperty>();
            foreach (SLMemberProperty mp in this.MemberProperties)
            {
                ph.MemberProperties.Add(mp.Clone());
            }

            ph.Members = new List<SLMembers>();
            foreach (SLMembers mems in this.Members)
            {
                ph.Members.Add(mems.Clone());
            }

            return ph;
        }
    }
}
