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
    internal class SLGroup
    {
        internal List<SLGroupMember> GroupMembers { get; set; }

        internal string Name { get; set; }
        internal string UniqueName { get; set; }
        internal string Caption { get; set; }
        internal string UniqueParent { get; set; }
        internal int? Id { get; set; }

        internal SLGroup()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.GroupMembers = new List<SLGroupMember>();

            this.Name = "";
            this.UniqueName = "";
            this.Caption = "";
            this.UniqueParent = "";
            this.Id = null;
        }

        internal void FromGroup(Group g)
        {
            this.SetAllNull();

            if (g.Name != null) this.Name = g.Name.Value;
            if (g.UniqueName != null) this.UniqueName = g.UniqueName.Value;
            if (g.Caption != null) this.Caption = g.Caption.Value;
            if (g.UniqueParent != null) this.UniqueParent = g.UniqueParent.Value;
            if (g.Id != null) this.Id = g.Id.Value;
        }

        internal Group ToGroup()
        {
            Group g = new Group();
            g.Name = this.Name;
            g.UniqueName = this.UniqueName;
            g.Caption = this.Caption;
            if (this.UniqueParent != null && this.UniqueParent.Length > 0) g.UniqueParent = this.UniqueParent;
            if (this.Id != null) g.Id = this.Id.Value;

            if (this.GroupMembers.Count > 0)
            {
                g.GroupMembers = new GroupMembers() { Count = (uint)this.GroupMembers.Count };
                foreach (SLGroupMember gm in this.GroupMembers)
                {
                    g.GroupMembers.Append(gm.ToGroupMember());
                }
            }

            return g;
        }

        internal SLGroup Clone()
        {
            SLGroup g = new SLGroup();
            g.Name = this.Name;
            g.UniqueName = this.UniqueName;
            g.Caption = this.Caption;
            g.UniqueParent = this.UniqueParent;
            g.Id = this.Id;

            g.GroupMembers = new List<SLGroupMember>();
            foreach (SLGroupMember gm in this.GroupMembers)
            {
                g.GroupMembers.Add(gm.Clone());
            }

            return g;
        }
    }
}
