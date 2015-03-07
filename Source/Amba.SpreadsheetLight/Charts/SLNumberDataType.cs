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
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Amba.SpreadsheetLight.Charts
{
    internal class SLNumberingCache : SLNumberDataType
    {
        internal SLNumberingCache() { }

        internal SLNumberingCache Clone()
        {
            SLNumberingCache nc = new SLNumberingCache();
            nc.FormatCode = this.FormatCode;
            nc.PointCount = this.PointCount;
            for (int i = 0; i < this.Points.Count; ++i)
            {
                nc.Points.Add(this.Points[i].Clone());
            }

            return nc;
        }
    }

    internal class SLNumberLiteral : SLNumberDataType
    {
        internal SLNumberLiteral() { }

        internal SLNumberLiteral Clone()
        {
            SLNumberLiteral nl = new SLNumberLiteral();
            nl.FormatCode = this.FormatCode;
            nl.PointCount = this.PointCount;
            for (int i = 0; i < this.Points.Count; ++i)
            {
                nl.Points.Add(this.Points[i].Clone());
            }

            return nl;
        }
    }

    /// <summary>
    /// For NumberingCache and NumberLiteral
    /// </summary>
    internal abstract class SLNumberDataType
    {
        internal string FormatCode { get; set; }
        internal uint PointCount { get; set; }
        internal List<SLNumericPoint> Points { get; set; }

        internal SLNumberDataType()
        {
            this.FormatCode = string.Empty;
            this.PointCount = 0;
            this.Points = new List<SLNumericPoint>();
        }

        internal C.NumberingCache ToNumberingCache()
        {
            C.NumberingCache nc = new C.NumberingCache();
            nc.FormatCode = new C.FormatCode(this.FormatCode);
            nc.PointCount = new C.PointCount() { Val = this.PointCount };
            for (int i = 0; i < this.Points.Count; ++i)
            {
                nc.Append(this.Points[i].ToNumericPoint());
            }

            return nc;
        }

        internal C.NumberLiteral ToNumberLiteral()
        {
            C.NumberLiteral nl = new C.NumberLiteral();
            nl.FormatCode = new C.FormatCode(this.FormatCode);
            nl.PointCount = new C.PointCount() { Val = this.PointCount };
            for (int i = 0; i < this.Points.Count; ++i)
            {
                nl.Append(this.Points[i].ToNumericPoint());
            }

            return nl;
        }
    }
}
