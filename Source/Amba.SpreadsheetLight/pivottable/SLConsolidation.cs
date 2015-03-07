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
    internal class SLConsolidation
    {
        internal List<List<string>> Pages { get; set; }
        internal List<SLRangeSet> RangeSets { get; set; }

        internal bool AutoPage { get; set; }

        internal SLConsolidation()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Pages = new List<List<string>>();
            this.RangeSets = new List<SLRangeSet>();
            this.AutoPage = true;
        }

        internal void FromConsolidation(Consolidation c)
        {
            this.SetAllNull();

            if (c.AutoPage != null) this.AutoPage = c.AutoPage.Value;

            Page pg;
            PageItem pgi;
            List<string> listPage;
            SLRangeSet rs;
            using (OpenXmlReader oxr = OpenXmlReader.Create(c))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Page))
                    {
                        listPage = new List<string>();
                        pg = (Page)oxr.LoadCurrentElement();
                        using (OpenXmlReader oxrPage = OpenXmlReader.Create(pg))
                        {
                            while (oxrPage.Read())
                            {
                                if (oxrPage.ElementType == typeof(PageItem))
                                {
                                    pgi = (PageItem)oxrPage.LoadCurrentElement();
                                    listPage.Add(pgi.Name.Value);
                                }
                            }
                        }
                        this.Pages.Add(listPage);
                    }
                    else if (oxr.ElementType == typeof(RangeSet))
                    {
                        rs = new SLRangeSet();
                        rs.FromRangeSet((RangeSet)oxr.LoadCurrentElement());
                        this.RangeSets.Add(rs);
                    }
                }
            }
        }

        internal Consolidation ToConsolidation()
        {
            Consolidation c = new Consolidation();
            if (this.AutoPage != true) c.AutoPage = this.AutoPage;

            if (this.Pages.Count > 0)
            {
                Page pg;
                c.Pages = new Pages() { Count = (uint)this.Pages.Count };
                foreach (List<string> ls in this.Pages)
                {
                    pg = new Page() { Count = (uint)ls.Count };
                    foreach (string s in ls)
                    {
                        pg.Append(new PageItem() { Name = s });
                    }
                    c.Pages.Append(pg);
                }
            }

            c.RangeSets = new RangeSets() { Count = (uint)this.RangeSets.Count };
            foreach (SLRangeSet rs in this.RangeSets)
            {
                c.RangeSets.Append(rs.ToRangeSet());
            }

            return c;
        }

        internal SLConsolidation Clone()
        {
            SLConsolidation c = new SLConsolidation();
            c.AutoPage = this.AutoPage;

            List<string> list;
            foreach (List<string> ls in this.Pages)
            {
                list = new List<string>();
                foreach (string s in ls)
                {
                    list.Add(s);
                }
                c.Pages.Add(list);
            }

            c.RangeSets = new List<SLRangeSet>();
            foreach (SLRangeSet rs in this.RangeSets)
            {
                c.RangeSets.Add(rs.Clone());
            }

            return c;
        }
    }
}
