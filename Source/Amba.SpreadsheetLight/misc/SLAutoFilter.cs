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
    internal class SLAutoFilter
    {
        internal List<SLFilterColumn> FilterColumns { get; set; }
        internal bool HasSortState;
        internal SLSortState SortState { get; set; }

        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        internal SLAutoFilter()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.FilterColumns = new List<SLFilterColumn>();
            this.SortState = new SLSortState();
            this.HasSortState = false;
            this.StartRowIndex = 1;
            this.StartColumnIndex = 1;
            this.EndRowIndex = 1;
            this.EndColumnIndex = 1;
        }

        internal void FromAutoFilter(AutoFilter af)
        {
            this.SetAllNull();

            int iStartRowIndex = 1;
            int iStartColumnIndex = 1;
            int iEndRowIndex = 1;
            int iEndColumnIndex = 1;
            string sRef = af.Reference.Value;
            if (sRef.IndexOf(":") > 0)
            {
                if (SLTool.FormatCellReferenceRangeToRowColumnIndex(sRef, out iStartRowIndex, out iStartColumnIndex, out iEndRowIndex, out iEndColumnIndex))
                {
                    this.StartRowIndex = iStartRowIndex;
                    this.StartColumnIndex = iStartColumnIndex;
                    this.EndRowIndex = iEndRowIndex;
                    this.EndColumnIndex = iEndColumnIndex;
                }
            }
            else
            {
                if (SLTool.FormatCellReferenceToRowColumnIndex(sRef, out iStartRowIndex, out iStartColumnIndex))
                {
                    this.StartRowIndex = iStartRowIndex;
                    this.StartColumnIndex = iStartColumnIndex;
                    this.EndRowIndex = iStartRowIndex;
                    this.EndColumnIndex = iStartColumnIndex;
                }
            }

            if (af.HasChildren)
            {
                SLFilterColumn fc;
                using (OpenXmlReader oxr = OpenXmlReader.Create(af))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(FilterColumn))
                        {
                            fc = new SLFilterColumn();
                            fc.FromFilterColumn((FilterColumn)oxr.LoadCurrentElement());
                            this.FilterColumns.Add(fc);
                        }
                        else if (oxr.ElementType == typeof(SortState))
                        {
                            this.SortState = new SLSortState();
                            this.SortState.FromSortState((SortState)oxr.LoadCurrentElement());
                            this.HasSortState = true;
                        }
                    }
                }
            }
        }

        internal AutoFilter ToAutoFilter()
        {
            AutoFilter af = new AutoFilter();

            if (this.StartRowIndex == this.EndRowIndex && this.StartColumnIndex == this.EndColumnIndex)
            {
                af.Reference = SLTool.ToCellReference(this.StartRowIndex, this.StartColumnIndex);
            }
            else
            {
                af.Reference = string.Format("{0}:{1}",
                    SLTool.ToCellReference(this.StartRowIndex, this.StartColumnIndex),
                    SLTool.ToCellReference(this.EndRowIndex, this.EndColumnIndex));
            }

            foreach (SLFilterColumn fc in this.FilterColumns)
            {
                af.Append(fc.ToFilterColumn());
            }

            if (HasSortState) af.Append(this.SortState.ToSortState());

            return af;
        }

        internal SLAutoFilter Clone()
        {
            SLAutoFilter af = new SLAutoFilter();
            af.FilterColumns = new List<SLFilterColumn>();
            for (int i = 0; i < this.FilterColumns.Count; ++i)
            {
                af.FilterColumns.Add(this.FilterColumns[i].Clone());
            }

            af.HasSortState = this.HasSortState;
            af.SortState = this.SortState.Clone();

            af.StartRowIndex = this.StartRowIndex;
            af.StartColumnIndex = this.StartColumnIndex;
            af.EndRowIndex = this.EndRowIndex;
            af.EndColumnIndex = this.EndColumnIndex;

            return af;
        }
    }
}
