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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Amba.SpreadsheetLight
{
    internal class SLCustomFilters
    {
        internal bool OneCustomFilter;

        internal bool HasFirstOperator;
        private FilterOperatorValues vFirstOperator;
        internal FilterOperatorValues FirstOperator
        {
            get { return vFirstOperator; }
            set
            {
                vFirstOperator = value;
                HasFirstOperator = vFirstOperator != FilterOperatorValues.Equal ? true : false;
            }
        }

        internal string FirstVal { get; set; }

        internal bool HasSecondOperator;
        private FilterOperatorValues vSecondOperator;
        internal FilterOperatorValues SecondOperator
        {
            get { return vSecondOperator; }
            set
            {
                vSecondOperator = value;
                HasSecondOperator = vSecondOperator != FilterOperatorValues.Equal ? true : false;
            }
        }

        internal string SecondVal { get; set; }

        internal bool? And { get; set; }

        internal SLCustomFilters()
        {
            CustomFilter cf = new CustomFilter();
            cf.Operator = FilterOperatorValues.Equal;
            cf.Val = "";
        }

        private void SetAllNull()
        {
            this.OneCustomFilter = true;
            this.vFirstOperator = FilterOperatorValues.Equal;
            this.HasFirstOperator = false;
            this.FirstVal = string.Empty;
            this.vSecondOperator = FilterOperatorValues.Equal;
            this.HasSecondOperator = false;
            this.SecondVal = string.Empty;
            this.And = null;
        }

        internal void FromCustomFilters(CustomFilters cfs)
        {
            this.SetAllNull();

            if (cfs.And != null && cfs.And.Value) this.And = cfs.And.Value;

            int i = 0;
            CustomFilter cf;
            using (OpenXmlReader oxr = OpenXmlReader.Create(cfs))
            {
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(CustomFilter))
                    {
                        ++i;
                        cf = (CustomFilter)oxr.LoadCurrentElement();
                        if (i == 1)
                        {
                            this.OneCustomFilter = true;
                            if (cf.Operator != null) this.FirstOperator = cf.Operator.Value;
                            if (cf.Val != null) this.FirstVal = cf.Val.Value;
                        }
                        else if (i == 2)
                        {
                            this.OneCustomFilter = false;
                            if (cf.Operator != null) this.SecondOperator = cf.Operator.Value;
                            if (cf.Val != null) this.SecondVal = cf.Val.Value;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
        }

        internal CustomFilters ToCustomFilters()
        {
            CustomFilters cfs = new CustomFilters();
            if (this.And != null && this.And.Value) cfs.And = this.And.Value;

            CustomFilter cf;
            if (this.OneCustomFilter)
            {
                cf = new CustomFilter();
                if (HasFirstOperator) cf.Operator = this.FirstOperator;
                cf.Val = this.FirstVal;
                cfs.Append(cf);
            }
            else
            {
                cf = new CustomFilter();
                if (HasFirstOperator) cf.Operator = this.FirstOperator;
                cf.Val = this.FirstVal;
                cfs.Append(cf);

                cf = new CustomFilter();
                if (HasSecondOperator) cf.Operator = this.SecondOperator;
                cf.Val = this.SecondVal;
                cfs.Append(cf);
            }

            return cfs;
        }

        internal SLCustomFilters Clone()
        {
            SLCustomFilters cfs = new SLCustomFilters();
            cfs.OneCustomFilter = this.OneCustomFilter;
            cfs.HasFirstOperator = this.HasFirstOperator;
            cfs.vFirstOperator = this.vFirstOperator;
            cfs.FirstVal = this.FirstVal;
            cfs.HasSecondOperator = this.HasSecondOperator;
            cfs.vSecondOperator = this.vSecondOperator;
            cfs.SecondVal = this.SecondVal;
            cfs.And = this.And;

            return cfs;
        }
    }
}
