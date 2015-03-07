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
    internal class SLPivotFilter
    {
        internal SLAutoFilter AutoFilter { get; set; }

        internal uint Field { get; set; }
        internal uint? MemberPropertyFieldId { get; set; }
        internal PivotFilterValues Type { get; set; }
        internal int EvaluationOrder { get; set; }
        internal uint Id { get; set; }
        internal uint? MeasureHierarchy { get; set; }
        internal uint? MeasureField { get; set; }
        internal string Name { get; set; }
        internal string Description { get; set; }
        internal string StringValue1 { get; set; }
        internal string StringValue2 { get; set; }

        internal SLPivotFilter()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.AutoFilter = new SLAutoFilter();

            this.Field = 0;
            this.MemberPropertyFieldId = null;
            this.Type = PivotFilterValues.Unknown;
            this.EvaluationOrder = 0;
            this.Id = 0;
            this.MeasureHierarchy = null;
            this.MeasureField = null;
            this.Name = "";
            this.Description = "";
            this.StringValue1 = "";
            this.StringValue2 = "";
        }

        internal void FromPivotFilter(PivotFilter pf)
        {
            this.SetAllNull();

            if (pf.Field != null) this.Field = pf.Field.Value;
            if (pf.MemberPropertyFieldId != null) this.MemberPropertyFieldId = pf.MemberPropertyFieldId.Value;
            if (pf.Type != null) this.Type = pf.Type.Value;
            if (pf.EvaluationOrder != null) this.EvaluationOrder = pf.EvaluationOrder.Value;
            if (pf.Id != null) this.Id = pf.Id.Value;
            if (pf.MeasureHierarchy != null) this.MeasureHierarchy = pf.MeasureHierarchy.Value;
            if (pf.MeasureField != null) this.MeasureField = pf.MeasureField.Value;
            if (pf.Name != null) this.Name = pf.Name.Value;
            if (pf.Description != null) this.Description = pf.Description.Value;
            if (pf.StringValue1 != null) this.StringValue1 = pf.StringValue1.Value;
            if (pf.StringValue2 != null) this.StringValue2 = pf.StringValue2.Value;

            if (pf.AutoFilter != null) this.AutoFilter.FromAutoFilter(pf.AutoFilter);
        }

        internal PivotFilter ToPivotFilter()
        {
            PivotFilter pf = new PivotFilter();
            pf.Field = this.Field;
            if (this.MemberPropertyFieldId != null) pf.MemberPropertyFieldId = this.MemberPropertyFieldId.Value;
            pf.Type = this.Type;
            if (this.EvaluationOrder != 0) pf.EvaluationOrder = this.EvaluationOrder;
            pf.Id = this.Id;
            if (this.MeasureHierarchy != null) pf.MeasureHierarchy = this.MeasureHierarchy.Value;
            if (this.MeasureField != null) pf.MeasureField = this.MeasureField.Value;
            if (this.Name != null && this.Name.Length > 0) pf.Name = this.Name;
            if (this.Description != null && this.Description.Length > 0) pf.Description = this.Description;
            if (this.StringValue1 != null && this.StringValue1.Length > 0) pf.StringValue1 = this.StringValue1;
            if (this.StringValue2 != null && this.StringValue2.Length > 0) pf.StringValue2 = this.StringValue2;

            pf.AutoFilter = this.AutoFilter.ToAutoFilter();

            return pf;
        }

        internal SLPivotFilter Clone()
        {
            SLPivotFilter pf = new SLPivotFilter();
            pf.Field = this.Field;
            pf.MemberPropertyFieldId = this.MemberPropertyFieldId;
            pf.Type = this.Type;
            pf.EvaluationOrder = this.EvaluationOrder;
            pf.Id = this.Id;
            pf.MeasureHierarchy = this.MeasureHierarchy;
            pf.MeasureField = this.MeasureField;
            pf.Name = this.Name;
            pf.Description = this.Description;
            pf.StringValue1 = this.StringValue1;
            pf.StringValue2 = this.StringValue2;

            pf.AutoFilter = this.AutoFilter.Clone();

            return pf;
        }
    }
}
