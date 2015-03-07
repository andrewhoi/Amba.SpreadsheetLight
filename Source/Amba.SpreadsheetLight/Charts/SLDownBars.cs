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
    /// Encapsulates properties and methods for down bars.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Charts.DownBars class.
    /// </summary>
    public class SLDownBars
    {
        internal SLA.SLShapeProperties ShapeProperties { get; set; }

        /// <summary>
        /// Fill properties.
        /// </summary>
        public SLA.SLFill Fill { get { return this.ShapeProperties.Fill; } }

        /// <summary>
        /// Border properties.
        /// </summary>
        public SLA.SLLinePropertiesType Border { get { return this.ShapeProperties.Outline; } }

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

        /// <summary>
        /// 3D format properties.
        /// </summary>
        public SLA.SLFormat3D Format3D { get { return this.ShapeProperties.Format3D; } }

        internal SLDownBars(List<System.Drawing.Color> ThemeColors, bool IsStylish = false)
        {
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors);
            if (IsStylish)
            {
                this.ShapeProperties.Fill.SetSolidFill(A.SchemeColorValues.Dark1, 0.35m, 0);
                this.ShapeProperties.Outline.Width = 0.75m;
                this.ShapeProperties.Outline.SetSolidLine(A.SchemeColorValues.Text1, 0.35m, 0);
            }
        }

        /// <summary>
        /// Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            this.ShapeProperties = new SLA.SLShapeProperties(this.ShapeProperties.listThemeColors);
        }

        internal C.DownBars ToDownBars(bool IsStylish = false)
        {
            C.DownBars db = new C.DownBars();

            if (this.ShapeProperties.HasShapeProperties) db.ChartShapeProperties = this.ShapeProperties.ToChartShapeProperties(IsStylish);

            return db;
        }

        internal SLDownBars Clone()
        {
            SLDownBars db = new SLDownBars(this.ShapeProperties.listThemeColors);
            db.ShapeProperties = this.ShapeProperties.Clone();

            return db;
        }
    }
}
