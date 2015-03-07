﻿// Copyright (c) 2011 Vincent Tan Wai Lip
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
    internal class SLItem
    {
        /// <summary>
        /// Attribute: n
        /// </summary>
        internal string ItemName { get; set; }

        /// <summary>
        /// Attribute: t
        /// </summary>
        internal ItemValues ItemType { get; set; }

        /// <summary>
        /// Attribute: h
        /// </summary>
        internal bool Hidden { get; set; }

        /// <summary>
        /// Attribute: s
        /// </summary>
        internal bool HasStringVlue { get; set; } // [sic]

        /// <summary>
        /// Attribute: sd
        /// </summary>
        internal bool HideDetails { get; set; }

        /// <summary>
        /// Attribute: f
        /// </summary>
        internal bool Calculated { get; set; }

        /// <summary>
        /// Attribute: m
        /// </summary>
        internal bool Missing { get; set; }

        /// <summary>
        /// Attribute: c
        /// </summary>
        internal bool ChildItems { get; set; }

        /// <summary>
        /// Attribute: x
        /// </summary>
        internal uint? Index { get; set; }
        
        /// <summary>
        /// Attribute: d
        /// </summary>
        internal bool Expanded { get; set; }

        /// <summary>
        /// Attribute: e
        /// </summary>
        internal bool DrillAcrossAttributes { get; set; }

        internal SLItem()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.ItemName = "";//n
            this.ItemType = ItemValues.Data;//t
            this.Hidden = false;//h
            this.HasStringVlue = false;//s
            this.HideDetails = true;//sd
            this.Calculated = false;//f
            this.Missing = false;//m
            this.ChildItems = false;//c
            this.Index = null;//uint opt x
            this.Expanded = false;//d
            this.DrillAcrossAttributes = true;//e
        }

        internal void FromItem(Item it)
        {
            this.SetAllNull();

            if (it.ItemName != null) this.ItemName = it.ItemName.Value;
            if (it.ItemType != null) this.ItemType = it.ItemType.Value;
            if (it.Hidden != null) this.Hidden = it.Hidden.Value;
            if (it.HasStringVlue != null) this.HasStringVlue = it.HasStringVlue.Value;
            if (it.HideDetails != null) this.HideDetails = it.HideDetails.Value;
            if (it.Calculated != null) this.Calculated = it.Calculated.Value;
            if (it.Missing != null) this.Missing = it.Missing.Value;
            if (it.ChildItems != null) this.ChildItems = it.ChildItems.Value;
            if (it.Index != null) this.Index = it.Index.Value;
            if (it.Expanded != null) this.Expanded = it.Expanded.Value;
            if (it.DrillAcrossAttributes != null) this.DrillAcrossAttributes = it.DrillAcrossAttributes.Value;
        }

        internal Item ToItem()
        {
            Item it = new Item();
            if (this.ItemName.Length > 0) it.ItemName = this.ItemName;
            if (this.ItemType != ItemValues.Data) it.ItemType = this.ItemType;
            if (this.Hidden != false) it.Hidden = this.Hidden;
            if (this.HasStringVlue != false) it.HasStringVlue = this.HasStringVlue;
            if (this.HideDetails != true) it.HideDetails = this.HideDetails;
            if (this.Calculated != false) it.Calculated = this.Calculated;
            if (this.Missing != false) it.Missing = this.Missing;
            if (this.ChildItems != false) it.ChildItems = this.ChildItems;
            if (this.Index != null) it.Index = this.Index.Value;
            if (this.Expanded != false) it.Expanded = this.Expanded;
            if (this.DrillAcrossAttributes != true) it.DrillAcrossAttributes = this.DrillAcrossAttributes;

            return it; // haha return it... maybe name a variable called "what"...
        }

        internal SLItem Clone()
        {
            SLItem it = new SLItem();
            it.ItemName = this.ItemName;
            it.ItemType = this.ItemType;
            it.Hidden = this.Hidden;
            it.HasStringVlue = this.HasStringVlue;
            it.HideDetails = this.HideDetails;
            it.Calculated = this.Calculated;
            it.Missing = this.Missing;
            it.ChildItems = this.ChildItems;
            it.Index = this.Index;
            it.Expanded = this.Expanded;
            it.DrillAcrossAttributes = this.DrillAcrossAttributes;

            return it;
        }
    }
}
