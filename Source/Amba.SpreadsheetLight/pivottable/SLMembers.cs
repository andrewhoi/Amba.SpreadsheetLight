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
    internal class SLMembers
    {
        internal List<string> Members { get; set; }
        internal uint? Level { get; set; }

        internal SLMembers()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Members = new List<string>();
            this.Level = null;
        }

        internal void FromMembers(Members m)
        {
            this.SetAllNull();

            if (m.Level != null) this.Level = m.Level.Value;

            Member mem;
            using (OpenXmlReader oxr = OpenXmlReader.Create(m))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(Member))
                    {
                        mem = (Member)oxr.LoadCurrentElement();
                        this.Members.Add(mem.Name.Value);
                    }
                }
            }
        }

        internal Members ToMembers()
        {
            Members m = new Members();
            m.Count = (uint)this.Members.Count;
            if (this.Level != null) m.Level = this.Level.Value;

            foreach (string s in this.Members)
            {
                m.Append(new Member() { Name = s });
            }

            return m;
        }

        internal SLMembers Clone()
        {
            SLMembers m = new SLMembers();
            m.Level = this.Level;

            m.Members = new List<string>();
            foreach (string s in this.Members)
            {
                m.Members.Add(s);
            }

            return m;
        }
    }
}
