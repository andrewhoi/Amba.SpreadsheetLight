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
    internal class SLMemberProperty
    {
        internal string Name { get; set; }
        internal bool ShowCell { get; set; }
        internal bool ShowTip { get; set; }
        internal bool ShowAsCaption { get; set; }
        internal uint? NameLength { get; set; }
        internal uint? PropertyNamePosition { get; set; }
        internal uint? PropertyNameLength { get; set; }
        internal uint? Level { get; set; }
        internal uint Field { get; set; }

        internal SLMemberProperty()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Name = "";
            this.ShowCell = false;
            this.ShowTip = false;
            this.ShowAsCaption = false;
            this.NameLength = null;
            this.PropertyNamePosition = null;
            this.PropertyNameLength = null;
            this.Level = null;
            this.Field = 0;
        }

        internal void FromMemberProperty(MemberProperty mp)
        {
            this.SetAllNull();

            if (mp.Name != null) this.Name = mp.Name.Value;
            if (mp.ShowCell != null) this.ShowCell = mp.ShowCell.Value;
            if (mp.ShowTip != null) this.ShowTip = mp.ShowTip.Value;
            if (mp.ShowAsCaption != null) this.ShowAsCaption = mp.ShowAsCaption.Value;
            if (mp.NameLength != null) this.NameLength = mp.NameLength.Value;
            if (mp.PropertyNamePosition != null) this.PropertyNamePosition = mp.PropertyNamePosition.Value;
            if (mp.PropertyNameLength != null) this.PropertyNameLength = mp.PropertyNameLength.Value;
            if (mp.Level != null) this.Level = mp.Level.Value;
            if (mp.Field != null) this.Field = mp.Field.Value;
        }

        internal MemberProperty ToMemberProperty()
        {
            MemberProperty mp = new MemberProperty();
            if (this.Name != null && this.Name.Length > 0) mp.Name = this.Name;
            if (this.ShowCell != false) mp.ShowCell = this.ShowCell;
            if (this.ShowTip != false) mp.ShowTip = this.ShowTip;
            if (this.ShowAsCaption != false) mp.ShowAsCaption = this.ShowAsCaption;
            if (this.NameLength != null) mp.NameLength = this.NameLength.Value;
            if (this.PropertyNamePosition != null) mp.PropertyNamePosition = this.PropertyNamePosition.Value;
            if (this.PropertyNameLength != null) mp.PropertyNameLength = this.PropertyNameLength.Value;
            if (this.Level != null) mp.Level = this.Level.Value;
            mp.Field = this.Field;

            return mp;
        }

        internal SLMemberProperty Clone()
        {
            SLMemberProperty mp = new SLMemberProperty();
            mp.Name = this.Name;
            mp.ShowCell = this.ShowCell;
            mp.ShowTip = this.ShowTip;
            mp.ShowAsCaption = this.ShowAsCaption;
            mp.NameLength = this.NameLength;
            mp.PropertyNamePosition = this.PropertyNamePosition;
            mp.PropertyNameLength = this.PropertyNameLength;
            mp.Level = this.Level;
            mp.Field = this.Field;

            return mp;
        }
    }
}
