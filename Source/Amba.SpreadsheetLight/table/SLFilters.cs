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
    internal class SLFilters
    {
        internal List<SLFilter> Filters { get; set; }
        internal List<SLDateGroupItem> DateGroupItems { get; set; }
        internal bool? Blank { get; set; }

        internal bool HasCalendarType;
        private CalendarValues vCalendarType;
        internal CalendarValues CalendarType
        {
            get { return vCalendarType; }
            set
            {
                vCalendarType = value;
                HasCalendarType = vCalendarType != CalendarValues.None ? true : false;
            }
        }

        internal SLFilters()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Filters = new List<SLFilter>();
            this.DateGroupItems = new List<SLDateGroupItem>();
            this.Blank = null;
            this.vCalendarType = CalendarValues.None;
            this.HasCalendarType = false;
        }

        internal void FromFilters(Filters fs)
        {
            this.SetAllNull();

            if (fs.Blank != null && fs.Blank.Value) this.Blank = fs.Blank.Value;
            if (fs.CalendarType != null) this.CalendarType = fs.CalendarType.Value;

            if (fs.HasChildren)
            {
                SLFilter f;
                SLDateGroupItem dgi;
                using (OpenXmlReader oxr = OpenXmlReader.Create(fs))
                {
                    while (oxr.Read())
                    {
                        if (oxr.ElementType == typeof(Filter))
                        {
                            f = new SLFilter();
                            f.FromFilter((Filter)oxr.LoadCurrentElement());
                            this.Filters.Add(f);
                        }
                        else if (oxr.ElementType == typeof(DateGroupItem))
                        {
                            dgi = new SLDateGroupItem();
                            dgi.FromDateGroupItem((DateGroupItem)oxr.LoadCurrentElement());
                            this.DateGroupItems.Add(dgi);
                        }
                    }
                }
            }
        }

        internal Filters ToFilters()
        {
            Filters fs = new Filters();
            if (this.Blank != null && this.Blank.Value) fs.Blank = this.Blank.Value;
            if (HasCalendarType) fs.CalendarType = this.CalendarType;

            foreach (SLFilter f in this.Filters)
            {
                fs.Append(f.ToFilter());
            }

            foreach (SLDateGroupItem dgi in this.DateGroupItems)
            {
                fs.Append(dgi.ToDateGroupItem());
            }

            return fs;
        }

        internal SLFilters Clone()
        {
            SLFilters fs = new SLFilters();

            int i;
            fs.Filters = new List<SLFilter>();
            for (i = 0; i < this.Filters.Count; ++i)
            {
                fs.Filters.Add(this.Filters[i].Clone());
            }

            fs.DateGroupItems = new List<SLDateGroupItem>();
            for (i = 0; i < this.DateGroupItems.Count; ++i)
            {
                fs.DateGroupItems.Add(this.DateGroupItems[i].Clone());
            }

            fs.Blank = this.Blank;
            fs.HasCalendarType = this.HasCalendarType;
            fs.vCalendarType = this.vCalendarType;

            return fs;
        }
    }
}
