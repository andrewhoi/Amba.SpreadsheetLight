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
using SLA = Amba.SpreadsheetLight.Drawing;

namespace Amba.SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for setting data markers in charts.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Charts.Marker class.
    /// </summary>
    public class SLMarker
    {
        internal bool HasMarker
        {
            get
            {
                return this.vSymbol != null || this.bySize != null || this.ShapeProperties.HasShapeProperties;
            }
        }

        internal C.MarkerStyleValues? vSymbol;
        /// <summary>
        /// Marker symbol.
        /// </summary>
        public C.MarkerStyleValues Symbol
        {
            get { return vSymbol ?? C.MarkerStyleValues.Auto; }
            set
            {
                vSymbol = value;
            }
        }

        internal byte? bySize;
        /// <summary>
        /// Range is 2 to 72 inclusive. Default is 5 in Open XML but Excel uses 7.
        /// </summary>
        public byte Size
        {
            get { return bySize ?? 5; }
            set
            {
                bySize = value;
                if (bySize != null)
                {
                    if (bySize < 2) bySize = 2;
                    if (bySize > 72) bySize = 72;
                }
            }
        }

        internal SLA.SLShapeProperties ShapeProperties { get; set; }

        /// <summary>
        /// Fill properties.
        /// </summary>
        public SLA.SLFill Fill { get { return this.ShapeProperties.Fill; } }

        /// <summary>
        /// Line properties.
        /// </summary>
        public SLA.SLLinePropertiesType Line { get { return this.ShapeProperties.Outline; } }

        internal SLMarker(List<System.Drawing.Color> ThemeColors)
        {
            this.vSymbol = null;
            this.bySize = null;
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors);
        }

        internal C.Marker ToMarker(bool IsStylish = false)
        {
            C.Marker m = new C.Marker();
            if (this.vSymbol != null) m.Symbol = new C.Symbol() { Val = this.vSymbol.Value };
            if (this.bySize != null) m.Size = new C.Size() { Val = this.bySize.Value };

            if (this.ShapeProperties.HasShapeProperties)
            {
                m.ChartShapeProperties = this.ShapeProperties.ToChartShapeProperties(IsStylish);
            }

            return m;
        }

        internal SLMarker Clone()
        {
            SLMarker m = new SLMarker(this.ShapeProperties.listThemeColors);
            m.Symbol = this.Symbol;
            m.bySize = this.bySize;
            m.ShapeProperties = this.ShapeProperties.Clone();

            return m;
        }
    }
}
