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
    internal class SLGroupLevel
    {
        internal List<SLGroup> Groups { get; set; }

        internal string UniqueName { get; set; }
        internal string Caption { get; set; }
        internal bool User { get; set; }
        internal bool CustomRollUp { get; set; }

        internal SLGroupLevel()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Groups = new List<SLGroup>();

            this.UniqueName = "";
            this.Caption = "";
            this.User = false;
            this.CustomRollUp = false;
        }

        internal void FromGroupLevel(GroupLevel gl)
        {
            this.SetAllNull();

            if (gl.UniqueName != null) this.UniqueName = gl.UniqueName.Value;
            if (gl.Caption != null) this.Caption = gl.Caption.Value;
            if (gl.User != null) this.User = gl.User.Value;
            if (gl.CustomRollUp != null) this.CustomRollUp = gl.CustomRollUp.Value;

            if (gl.Groups != null)
            {
                SLGroup g;
                using (OpenXmlReader oxr = OpenXmlReader.Create(gl.Groups))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(Group))
                        {
                            g = new SLGroup();
                            g.FromGroup((Group)oxr.LoadCurrentElement());
                            this.Groups.Add(g);
                        }
                    }
                }
            }
        }

        internal GroupLevel ToGroupLevel()
        {
            GroupLevel gl = new GroupLevel();
            gl.UniqueName = this.UniqueName;
            gl.Caption = this.Caption;
            if (this.User != false) gl.User = this.User;
            if (this.CustomRollUp != false) gl.CustomRollUp = this.CustomRollUp;

            if (this.Groups.Count > 0)
            {
                gl.Groups = new Groups() { Count = (uint)this.Groups.Count };
                foreach (SLGroup g in this.Groups)
                {
                    gl.Groups.Append(g.ToGroup());
                }
            }

            return gl;
        }

        internal SLGroupLevel Clone()
        {
            SLGroupLevel gl = new SLGroupLevel();
            gl.UniqueName = this.UniqueName;
            gl.Caption = this.Caption;
            gl.User = this.User;
            gl.CustomRollUp = this.CustomRollUp;

            gl.Groups = new List<SLGroup>();
            foreach (SLGroup g in this.Groups)
            {
                gl.Groups.Add(g.Clone());
            }

            return gl;
        }
    }
}
