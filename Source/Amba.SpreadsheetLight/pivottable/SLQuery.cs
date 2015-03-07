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
using DocumentFormat.OpenXml.Spreadsheet;

namespace Amba.SpreadsheetLight
{
    internal class SLQuery
    {
        internal bool HasTuples;
        internal SLTuplesType Tuples { get; set; }

        internal string Mdx { get; set; }

        internal SLQuery()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.HasTuples = false;
            this.Tuples = new SLTuplesType();

            this.Mdx = "";
        }

        internal void FromQuery(Query q)
        {
            this.SetAllNull();

            if (q.Mdx != null) this.Mdx = q.Mdx.Value;

            if (q.Tuples != null)
            {
                this.Tuples.FromTuples(q.Tuples);
                this.HasTuples = true;
            }
        }

        internal Query ToQuery()
        {
            Query q = new Query();
            q.Mdx = this.Mdx;

            if (this.HasTuples) q.Tuples = this.Tuples.ToTuples();

            return q;
        }

        internal SLQuery Clone()
        {
            SLQuery q = new SLQuery();
            q.Mdx = this.Mdx;
            q.HasTuples = this.HasTuples;
            q.Tuples = this.Tuples.Clone();

            return q;
        }
    }
}
