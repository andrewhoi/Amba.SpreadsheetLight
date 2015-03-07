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
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

// Apparently .NET Framework 4 has a System.Tuple, which clashes
// with DocumentFormat.OpenXml.Spreadsheet.Tuple.
// Good thing we're on 3.5...

namespace Amba.SpreadsheetLight
{
    /// <summary>
    /// This doubles for SortByTuple and Tuples
    /// </summary>
    internal class SLTuplesType
    {
        internal List<SLTuple> Tuples { get; set; }
        internal uint? MemberNameCount { get; set; }

        internal SLTuplesType()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Tuples = new List<SLTuple>();
            this.MemberNameCount = null;
        }

        internal void FromSortByTuple(SortByTuple sbt)
        {
            this.SetAllNull();

            if (sbt.MemberNameCount != null) this.MemberNameCount = sbt.MemberNameCount.Value;

            SLTuple t;
            using (OpenXmlReader oxr = OpenXmlReader.Create(sbt))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Tuple))
                    {
                        t = new SLTuple();
                        t.FromTuple((Tuple)oxr.LoadCurrentElement());
                        this.Tuples.Add(t);
                    }
                }
            }
        }

        internal void FromTuples(Tuples tpls)
        {
            this.SetAllNull();

            if (tpls.MemberNameCount != null) this.MemberNameCount = tpls.MemberNameCount.Value;

            SLTuple t;
            using (OpenXmlReader oxr = OpenXmlReader.Create(tpls))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Tuple))
                    {
                        t = new SLTuple();
                        t.FromTuple((Tuple)oxr.LoadCurrentElement());
                        this.Tuples.Add(t);
                    }
                }
            }
        }

        internal SortByTuple ToSortByTuple()
        {
            SortByTuple sbt = new SortByTuple();
            if (this.MemberNameCount != null) sbt.MemberNameCount = this.MemberNameCount.Value;

            foreach (SLTuple t in this.Tuples)
            {
                sbt.Append(t.ToTuple());
            }

            return sbt;
        }

        internal Tuples ToTuples()
        {
            Tuples tpls = new Tuples();
            if (this.MemberNameCount != null) tpls.MemberNameCount = this.MemberNameCount.Value;

            foreach (SLTuple t in this.Tuples)
            {
                tpls.Append(t.ToTuple());
            }

            return tpls;
        }

        internal SLTuplesType Clone()
        {
            SLTuplesType tt = new SLTuplesType();
            tt.MemberNameCount = this.MemberNameCount;

            tt.Tuples = new List<SLTuple>();
            foreach (SLTuple t in this.Tuples)
            {
                tt.Tuples.Add(t.Clone());
            }

            return tt;
        }
    }
}
