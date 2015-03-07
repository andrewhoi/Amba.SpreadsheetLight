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
    internal class SLTupleCache
    {
        internal bool HasEntries;
        internal SLEntries Entries { get; set; }
        internal List<SLTupleSet> Sets { get; set; }
        internal List<SLQuery> QueryCache { get; set; }
        internal List<SLServerFormat> ServerFormats { get; set; }

        internal SLTupleCache()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.HasEntries = false;
            this.Entries = new SLEntries();
            this.Sets = new List<SLTupleSet>();
            this.QueryCache = new List<SLQuery>();
            this.ServerFormats = new List<SLServerFormat>();
        }

        internal void FromTupleCache(TupleCache tc)
        {
            this.SetAllNull();

            // I decided to do this one by one instead of just running through all the child
            // elements. Mainly because this seems safer... so complicated! It's just a pivot table
            // for goodness sakes...

            if (tc.Entries != null)
            {
                this.Entries.FromEntries(tc.Entries);
                this.HasEntries = true;
            }

            if (tc.Sets != null)
            {
                SLTupleSet ts;
                using (OpenXmlReader oxr = OpenXmlReader.Create(tc.Sets))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(TupleSet))
                        {
                            ts = new SLTupleSet();
                            ts.FromTupleSet((TupleSet)oxr.LoadCurrentElement());
                            this.Sets.Add(ts);
                        }
                    }
                }
            }

            if (tc.QueryCache != null)
            {
                SLQuery q;
                using (OpenXmlReader oxr = OpenXmlReader.Create(tc.QueryCache))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(Query))
                        {
                            q = new SLQuery();
                            q.FromQuery((Query)oxr.LoadCurrentElement());
                            this.QueryCache.Add(q);
                        }
                    }
                }
            }

            if (tc.ServerFormats != null)
            {
                SLServerFormat sf;
                using (OpenXmlReader oxr = OpenXmlReader.Create(tc.ServerFormats))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(ServerFormat))
                        {
                            sf = new SLServerFormat();
                            sf.FromServerFormat((ServerFormat)oxr.LoadCurrentElement());
                            this.ServerFormats.Add(sf);
                        }
                    }
                }
            }
        }

        internal TupleCache ToTupleCache()
        {
            TupleCache tc = new TupleCache();
            if (this.HasEntries) tc.Entries = this.Entries.ToEntries();

            if (this.Sets.Count > 0)
            {
                tc.Sets = new Sets() { Count = (uint)this.Sets.Count };
                foreach (SLTupleSet ts in this.Sets)
                {
                    tc.Sets.Append(ts.ToTupleSet());
                }
            }

            if (this.QueryCache.Count > 0)
            {
                tc.QueryCache = new QueryCache() { Count = (uint)this.QueryCache.Count };
                foreach (SLQuery q in this.QueryCache)
                {
                    tc.QueryCache.Append(q.ToQuery());
                }
            }

            if (this.ServerFormats.Count > 0)
            {
                tc.ServerFormats = new ServerFormats() { Count = (uint)this.ServerFormats.Count };
                foreach (SLServerFormat sf in this.ServerFormats)
                {
                    tc.ServerFormats.Append(sf.ToServerFormat());
                }
            }

            return tc;
        }

        internal SLTupleCache Clone()
        {
            SLTupleCache tc = new SLTupleCache();
            tc.HasEntries = this.HasEntries;
            tc.Entries = this.Entries.Clone();

            tc.Sets = new List<SLTupleSet>();
            foreach (SLTupleSet ts in this.Sets)
            {
                tc.Sets.Add(ts.Clone());
            }

            tc.QueryCache = new List<SLQuery>();
            foreach (SLQuery q in this.QueryCache)
            {
                tc.QueryCache.Add(q.Clone());
            }

            tc.ServerFormats = new List<SLServerFormat>();
            foreach (SLServerFormat sf in this.ServerFormats)
            {
                tc.ServerFormats.Add(sf.Clone());
            }

            return tc;
        }
    }
}
