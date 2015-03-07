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

namespace Amba.SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for setting data label options for charts.
    /// </summary>
    public class SLDataLabelOptions : EGDLblShared
    {
        // TODO Layout?

        internal SLRstType RichText { get; set; }

        internal SLDataLabelOptions(List<System.Drawing.Color> ThemeColors) : base(ThemeColors)
        {
            this.RichText = null;
        }

        /// <summary>
        /// Set custom label text.
        /// </summary>
        /// <param name="Text">The custom text.</param>
        public void SetLabelText(string Text)
        {
            SLRstType rst = new SLRstType();
            rst.AppendText(Text);
            this.RichText = rst.Clone();
        }

        /// <summary>
        /// Set custom label text.
        /// </summary>
        /// <param name="RichText">The custom text in rich text format.</param>
        public void SetLabelText(SLRstType RichText)
        {
            this.RichText = RichText.Clone();
        }

        /// <summary>
        /// Reset the label text. This removes any custom label text.
        /// </summary>
        public void ResetLabelText()
        {
            this.RichText = null;
        }

        internal C.DataLabel ToDataLabel(int index)
        {
            C.DataLabel lbl = new C.DataLabel();

            lbl.Index = new C.Index() { Val = (uint)index };

            lbl.Append(new C.Layout());

            if (this.RichText != null || this.Rotation != null || this.Vertical != null || this.Anchor != null || this.AnchorCenter != null)
            {
                C.ChartText ctxt = new C.ChartText();
                ctxt.RichText = new C.RichText();
                ctxt.RichText.BodyProperties = new A.BodyProperties();

                if (this.Rotation != null || this.Vertical != null || this.Anchor != null || this.AnchorCenter != null)
                {
                    if (this.Rotation != null) ctxt.RichText.BodyProperties.Rotation = (int)(this.Rotation.Value * SLConstants.DegreeToAngleRepresentation);
                    if (this.Vertical != null) ctxt.RichText.BodyProperties.Vertical = this.Vertical.Value;
                    if (this.Anchor != null) ctxt.RichText.BodyProperties.Anchor = this.Anchor.Value;
                    if (this.AnchorCenter != null) ctxt.RichText.BodyProperties.AnchorCenter = this.AnchorCenter.Value;
                }

                ctxt.RichText.ListStyle = new A.ListStyle();

                if (this.RichText != null) ctxt.RichText.Append(this.RichText.ToParagraph());

                lbl.Append(ctxt);
            }

            if (this.HasNumberingFormat)
            {
                lbl.Append(new C.NumberingFormat() { FormatCode = this.FormatCode, SourceLinked = this.SourceLinked });
            }

            if (this.ShapeProperties.HasShapeProperties) lbl.Append(this.ShapeProperties.ToChartShapeProperties());

            if (this.vLabelPosition != null) lbl.Append(new C.DataLabelPosition() { Val = this.vLabelPosition.Value });

            lbl.Append(new C.ShowLegendKey() { Val = this.ShowLegendKey });
            lbl.Append(new C.ShowValue() { Val = this.ShowValue });
            lbl.Append(new C.ShowCategoryName() { Val = this.ShowCategoryName });
            lbl.Append(new C.ShowSeriesName() { Val = this.ShowSeriesName });
            lbl.Append(new C.ShowPercent() { Val = this.ShowPercentage });
            lbl.Append(new C.ShowBubbleSize() { Val = this.ShowBubbleSize });

            if (this.Separator != null && this.Separator.Length > 0) lbl.Append(new C.Separator() { Text = this.Separator });

            return lbl;
        }

        internal SLDataLabelOptions Clone()
        {
            SLDataLabelOptions dlo = new SLDataLabelOptions(this.ShapeProperties.listThemeColors);
            dlo.Rotation = this.Rotation;
            dlo.Vertical = this.Vertical;
            dlo.Anchor = this.Anchor;
            dlo.AnchorCenter = this.AnchorCenter;
            dlo.HasNumberingFormat = this.HasNumberingFormat;
            dlo.sFormatCode = this.sFormatCode;
            dlo.bSourceLinked = this.bSourceLinked;
            dlo.vLabelPosition = this.vLabelPosition;
            dlo.ShapeProperties = this.ShapeProperties.Clone();
            dlo.ShowLegendKey = this.ShowLegendKey;
            dlo.ShowValue = this.ShowValue;
            dlo.ShowCategoryName = this.ShowCategoryName;
            dlo.ShowSeriesName = this.ShowSeriesName;
            dlo.ShowPercentage = this.ShowPercentage;
            dlo.ShowBubbleSize = this.ShowBubbleSize;
            dlo.Separator = this.Separator;
            if (this.RichText != null) dlo.RichText = this.RichText.Clone();

            return dlo;
        }
    }
}
