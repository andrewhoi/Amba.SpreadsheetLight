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
using DocumentFormat.OpenXml.Spreadsheet;

namespace Amba.SpreadsheetLight
{
    internal class SLCalculatedMember
    {
        internal string Name { get; set; }
        internal string Mdx { get; set; }
        internal string MemberName { get; set; }
        internal string Hierarchy { get; set; }
        internal string ParentName { get; set; }
        internal int SolveOrder { get; set; }
        internal bool Set { get; set; }

        internal SLCalculatedMember()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Name = "";
            this.Mdx = "";
            this.MemberName = "";
            this.Hierarchy = "";
            this.ParentName = "";
            this.SolveOrder = 0;
            this.Set = false;
        }

        internal void FromCalculatedMember(CalculatedMember cm)
        {
            this.SetAllNull();

            if (cm.Name != null) this.Name = cm.Name.Value;
            if (cm.Mdx != null) this.Mdx = cm.Mdx.Value;
            if (cm.MemberName != null) this.MemberName = cm.MemberName.Value;
            if (cm.Hierarchy != null) this.Hierarchy = cm.Hierarchy.Value;
            if (cm.ParentName != null) this.ParentName = cm.ParentName.Value;
            if (cm.SolveOrder != null) this.SolveOrder = cm.SolveOrder.Value;
            if (cm.Set != null) this.Set = cm.Set.Value;
        }

        internal CalculatedMember ToCalculatedMember()
        {
            CalculatedMember cm = new CalculatedMember();
            cm.Name = this.Name;
            cm.Mdx = this.Mdx;
            if (this.MemberName != null && this.MemberName.Length > 0) cm.MemberName = this.MemberName;
            if (this.Hierarchy != null && this.Hierarchy.Length > 0) cm.Hierarchy = this.Hierarchy;
            if (this.ParentName != null && this.ParentName.Length > 0) cm.ParentName = this.ParentName;
            if (this.SolveOrder != 0) cm.SolveOrder = this.SolveOrder;
            if (this.Set != false) cm.Set = this.Set;

            return cm;
        }

        internal SLCalculatedMember Clone()
        {
            SLCalculatedMember cm = new SLCalculatedMember();
            cm.Name = this.Name;
            cm.Mdx = this.Mdx;
            cm.MemberName = this.MemberName;
            cm.Hierarchy = this.Hierarchy;
            cm.ParentName = this.ParentName;
            cm.SolveOrder = this.SolveOrder;
            cm.Set = this.Set;

            return cm;
        }
    }
}
