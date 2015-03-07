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
    internal class SLTupleSet
    {
        //CT_Set

        internal List<SLTuplesType> Tuples { get; set; }

        internal bool HasSortByTuple;
        internal SLTuplesType SortByTuple { get; set; }

        // count is for number of Tuples
        internal int MaxRank { get; set; }
        internal string SetDefinition { get; set; }
        internal SortValues SortType { get; set; }
        internal bool QueryFailed { get; set; }

        internal SLTupleSet()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Tuples = new List<SLTuplesType>();

            this.HasSortByTuple = false;
            this.SortByTuple = new SLTuplesType();

            this.MaxRank = 0;
            this.SetDefinition = "";
            this.SortType = SortValues.None;
            this.QueryFailed = false;
        }

        internal void FromTupleSet(TupleSet ts)
        {
            this.SetAllNull();

            if (ts.MaxRank != null) this.MaxRank = ts.MaxRank.Value;
            if (ts.SetDefinition != null) this.SetDefinition = ts.SetDefinition.Value;
            if (ts.SortType != null) this.SortType = ts.SortType.Value;
            if (ts.QueryFailed != null) this.QueryFailed = ts.QueryFailed.Value;

            SLTuplesType tt;
            using (OpenXmlReader oxr = OpenXmlReader.Create(ts))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Tuples))
                    {
                        tt = new SLTuplesType();
                        tt.FromTuples((Tuples)oxr.LoadCurrentElement());
                        this.Tuples.Add(tt);
                    }
                    else if (oxr.ElementType == typeof(SortByTuple))
                    {
                        this.SortByTuple.FromSortByTuple((SortByTuple)oxr.LoadCurrentElement());
                        this.HasSortByTuple = true;
                    }
                }
            }
        }

        internal TupleSet ToTupleSet()
        {
            TupleSet ts = new TupleSet();
            if (this.Tuples.Count > 0) ts.Count = (uint)this.Tuples.Count;
            ts.MaxRank = this.MaxRank;
            ts.SetDefinition = this.SetDefinition;
            if (this.SortType != SortValues.None) ts.SortType = this.SortType;
            if (this.QueryFailed != false) ts.QueryFailed = this.QueryFailed;

            if (this.Tuples.Count > 0)
            {
                foreach (SLTuplesType tt in this.Tuples)
                {
                    ts.Append(tt.ToTuples());
                }
            }

            if (this.HasSortByTuple)
            {
                ts.Append(this.SortByTuple.ToSortByTuple());
            }

            return ts;
        }

        internal SLTupleSet Clone()
        {
            SLTupleSet ts = new SLTupleSet();
            ts.MaxRank = this.MaxRank;
            ts.SetDefinition = this.SetDefinition;
            ts.SortType = this.SortType;
            ts.QueryFailed = this.QueryFailed;

            ts.Tuples = new List<SLTuplesType>();
            foreach (SLTuplesType tt in this.Tuples)
            {
                ts.Tuples.Add(tt.Clone());
            }

            ts.HasSortByTuple = this.HasSortByTuple;
            ts.SortByTuple = this.SortByTuple.Clone();

            return ts;
        }
    }
}
