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
    /// Encapsulates properties and methods for setting titles for charts.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Charts.Title class.
    /// </summary>
    public class SLTitle : SLChartAlignment
    {
        internal SLRstType rst { get; set; }

        /// <summary>
        /// Title text. This returns the plain text version if rich text is applied.
        /// </summary>
        public string Text
        {
            get { return this.rst.ToPlainString(); }
            set
            {
                this.rst = new SLRstType(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
                this.rst.SetText(value);
            }
        }

        /// <summary>
        /// Specifies if the title overlaps.
        /// </summary>
        public bool Overlay { get; set; }

        internal SLA.SLShapeProperties ShapeProperties;

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

        internal SLTitle(List<System.Drawing.Color> ThemeColors, bool IsStylish = false)
        {
            // just put in the theme colors, even though it's probably not needed.
            // Memory optimisations? Take it out.
            this.rst = new SLRstType(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, ThemeColors, new List<System.Drawing.Color>());
            this.Overlay = false;
            this.ShapeProperties = new SLA.SLShapeProperties(ThemeColors);

            if (IsStylish)
            {
                this.ShapeProperties.Fill.SetNoFill();
                this.ShapeProperties.Outline.SetNoLine();
            }

            this.RemoveTextAlignment();
        }

        /// <summary>
        /// Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            this.ShapeProperties = new SLA.SLShapeProperties(this.ShapeProperties.listThemeColors);
        }

        /// <summary>
        /// Set the title text.
        /// </summary>
        /// <param name="Text">The title text.</param>
        public void SetTitle(string Text)
        {
            this.Text = Text;
        }

        /// <summary>
        /// Set the title with a rich text string.
        /// </summary>
        /// <param name="RichText">The rich text.</param>
        public void SetTitle(SLRstType RichText)
        {
            this.rst = RichText.Clone();
        }

        internal C.Title ToTitle(bool IsStylish = false)
        {
            C.Title t = new C.Title();

            bool bHasText = this.rst.ToPlainString().Length > 0;
            if (bHasText || this.Rotation != null || this.Vertical != null || this.Anchor != null || this.AnchorCenter != null)
            {
                t.ChartText = new C.ChartText();
                t.ChartText.RichText = new C.RichText();
                t.ChartText.RichText.BodyProperties = new A.BodyProperties();

                if (this.Rotation != null || this.Vertical != null || this.Anchor != null || this.AnchorCenter != null)
                {
                    if (this.Rotation != null) t.ChartText.RichText.BodyProperties.Rotation = (int)(this.Rotation.Value * SLConstants.DegreeToAngleRepresentation);
                    if (this.Vertical != null) t.ChartText.RichText.BodyProperties.Vertical = this.Vertical.Value;
                    if (this.Anchor != null) t.ChartText.RichText.BodyProperties.Anchor = this.Anchor.Value;
                    if (this.AnchorCenter != null) t.ChartText.RichText.BodyProperties.AnchorCenter = this.AnchorCenter.Value;
                }

                t.ChartText.RichText.ListStyle = new A.ListStyle();

                if (bHasText) t.ChartText.RichText.Append(this.rst.ToParagraph());
            }

            t.Layout = new C.Layout();
            t.Overlay = new C.Overlay() { Val = this.Overlay };
            if (this.ShapeProperties.HasShapeProperties) t.ChartShapeProperties = this.ShapeProperties.ToChartShapeProperties(IsStylish);

            return t;
        }

        internal SLTitle Clone()
        {
            SLTitle t = new SLTitle(this.ShapeProperties.listThemeColors);
            t.Rotation = this.Rotation;
            t.Vertical = this.Vertical;
            t.Anchor = this.Anchor;
            t.AnchorCenter = this.AnchorCenter;
            t.rst = this.rst.Clone();
            t.Overlay = this.Overlay;
            t.ShapeProperties = this.ShapeProperties.Clone();

            return t;
        }
    }
}
