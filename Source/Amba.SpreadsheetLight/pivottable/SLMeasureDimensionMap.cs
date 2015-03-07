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
    internal class SLMeasureDimensionMap
    {
        internal uint? MeasureGroup { get; set; }
        internal uint? Dimension { get; set; }

        internal SLMeasureDimensionMap()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.MeasureGroup = null;
            this.Dimension = null;
        }

        internal void FromMeasureDimensionMap(MeasureDimensionMap mdm)
        {
            this.SetAllNull();

            if (mdm.MeasureGroup != null) this.MeasureGroup = mdm.MeasureGroup.Value;
            if (mdm.Dimension != null) this.Dimension = mdm.Dimension.Value;
        }

        internal MeasureDimensionMap ToMeasureDimensionMap()
        {
            MeasureDimensionMap mdm = new MeasureDimensionMap();
            if (this.MeasureGroup != null) mdm.MeasureGroup = this.MeasureGroup.Value;
            if (this.Dimension != null) mdm.Dimension = this.Dimension.Value;

            return mdm;
        }

        internal SLMeasureDimensionMap Clone()
        {
            SLMeasureDimensionMap mdm = new SLMeasureDimensionMap();
            mdm.MeasureGroup = this.MeasureGroup;
            mdm.Dimension = this.Dimension;

            return mdm;
        }
    }
}
