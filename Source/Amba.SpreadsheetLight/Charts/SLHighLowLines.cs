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
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLA = Amba.SpreadsheetLight.Drawing;

namespace Amba.SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for high-low lines.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Charts.HighLowLines class.
    /// </summary>
    public class SLHighLowLines
    {
        internal SLA.SLShapeProperties ShapeProperties { get; set; }

        /// <summary>
        /// Line properties.
        /// </summary>
        public SLA.SLLinePropertiesType Line { get { return this.ShapeProperties.Outline; } }

        /// <summary>
        /// Shadow properties.
        /// </summary>
        public SLA.SLShadowEffect Shadow { get { return this.ShapeProperties.EffectList.Shadow; } }

        /// <summary>
        /// Glow properties.
        /// </summary>
        public SLA.SLGlow Glow { get { return this.ShapeProperties.EffectList.Glow; } }

        /// <summary>
        /// Soft edge properties.
        /// </summary>
        public SLA.SLSoftEdge SoftEdge { get { return this.ShapeProperties.EffectList.SoftEdge; } }

        internal SLHighLowLines(List<System.Drawing.Color> ThemeColors, bool IsStylish = false)
        {
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors);
            if (IsStylish)
            {
                this.ShapeProperties.Outline.Width = 0.75m;
                this.ShapeProperties.Outline.CapType = A.LineCapValues.Flat;
                this.ShapeProperties.Outline.CompoundLineType = A.CompoundLineValues.Single;
                this.ShapeProperties.Outline.Alignment = A.PenAlignmentValues.Center;
                this.ShapeProperties.Outline.SetSolidLine(A.SchemeColorValues.Text1, 0.25m, 0);
                this.ShapeProperties.Outline.JoinType = SLA.SLLineJoinValues.Round;
            }
        }

        /// <summary>
        /// Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            this.ShapeProperties = new SLA.SLShapeProperties(this.ShapeProperties.listThemeColors);
        }

        internal C.HighLowLines ToHighLowLines(bool IsStylish = false)
        {
            C.HighLowLines hll = new C.HighLowLines();

            if (this.ShapeProperties.HasShapeProperties) hll.ChartShapeProperties = this.ShapeProperties.ToChartShapeProperties(IsStylish);

            return hll;
        }

        internal SLHighLowLines Clone()
        {
            SLHighLowLines hll = new SLHighLowLines(this.ShapeProperties.listThemeColors);
            hll.ShapeProperties = this.ShapeProperties.Clone();

            return hll;
        }
    }
}
