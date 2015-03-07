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
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Amba.SpreadsheetLight.Charts
{
    /// <summary>
    /// Encapsulates properties and methods for setting group data label options for charts.
    /// </summary>
    public class SLGroupDataLabelOptions : EGDLblShared
    {
        /// <summary>
        /// Specifies if leader lines are shown. This is for pie charts (I think...).
        /// </summary>
        public bool ShowLeaderLines { get; set; }
        // TODO Leaderlines (pie charts)

        internal SLGroupDataLabelOptions(List<System.Drawing.Color> ThemeColors) : base(ThemeColors)
        {
            this.ShowLeaderLines = false;
        }

        internal C.DataLabels ToDataLabels(Dictionary<int, SLDataLabelOptions> Options, bool ToDelete)
        {
            C.DataLabels lbls = new C.DataLabels();

            if (Options.Count > 0)
            {
                List<int> indexlist = Options.Keys.ToList<int>();
                indexlist.Sort();
                int index;
                for (int i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    lbls.Append(Options[index].ToDataLabel(index));
                }
            }

            if (ToDelete)
            {
                lbls.Append(new C.Delete() { Val = true });
            }
            else
            {
                if (this.HasNumberingFormat)
                {
                    lbls.Append(new C.NumberingFormat() { FormatCode = this.FormatCode, SourceLinked = this.SourceLinked });
                }

                if (this.ShapeProperties.HasShapeProperties) lbls.Append(this.ShapeProperties.ToChartShapeProperties());

                if (this.Rotation != null || this.Vertical != null || this.Anchor != null || this.AnchorCenter != null)
                {
                    C.TextProperties txtprops = new C.TextProperties();
                    txtprops.BodyProperties = new A.BodyProperties();
                    if (this.Rotation != null) txtprops.BodyProperties.Rotation = (int)(this.Rotation.Value * SLConstants.DegreeToAngleRepresentation);
                    if (this.Vertical != null) txtprops.BodyProperties.Vertical = this.Vertical.Value;
                    if (this.Anchor != null) txtprops.BodyProperties.Anchor = this.Anchor.Value;
                    if (this.AnchorCenter != null) txtprops.BodyProperties.AnchorCenter = this.AnchorCenter.Value;

                    txtprops.ListStyle = new A.ListStyle();

                    A.Paragraph para = new A.Paragraph();
                    para.ParagraphProperties = new A.ParagraphProperties();
                    para.ParagraphProperties.Append(new A.DefaultRunProperties());
                    txtprops.Append(para);

                    lbls.Append(txtprops);
                }

                if (this.vLabelPosition != null) lbls.Append(new C.DataLabelPosition() { Val = this.vLabelPosition.Value });

                lbls.Append(new C.ShowLegendKey() { Val = this.ShowLegendKey });
                lbls.Append(new C.ShowValue() { Val = this.ShowValue });
                lbls.Append(new C.ShowCategoryName() { Val = this.ShowCategoryName });
                lbls.Append(new C.ShowSeriesName() { Val = this.ShowSeriesName });
                lbls.Append(new C.ShowPercent() { Val = this.ShowPercentage });
                lbls.Append(new C.ShowBubbleSize() { Val = this.ShowBubbleSize });

                if (this.Separator != null && this.Separator.Length > 0) lbls.Append(new C.Separator() { Text = this.Separator });

                if (this.ShowLeaderLines) lbls.Append(new C.ShowLeaderLines() { Val = this.ShowLeaderLines });
            }

            return lbls;
        }

        internal SLGroupDataLabelOptions Clone()
        {
            SLGroupDataLabelOptions gdlo = new SLGroupDataLabelOptions(this.ShapeProperties.listThemeColors);
            gdlo.Rotation = this.Rotation;
            gdlo.Vertical = this.Vertical;
            gdlo.Anchor = this.Anchor;
            gdlo.AnchorCenter = this.AnchorCenter;
            gdlo.HasNumberingFormat = this.HasNumberingFormat;
            gdlo.sFormatCode = this.sFormatCode;
            gdlo.bSourceLinked = this.bSourceLinked;
            gdlo.vLabelPosition = this.vLabelPosition;
            gdlo.ShapeProperties = this.ShapeProperties.Clone();
            gdlo.ShowLegendKey = this.ShowLegendKey;
            gdlo.ShowValue = this.ShowValue;
            gdlo.ShowCategoryName = this.ShowCategoryName;
            gdlo.ShowSeriesName = this.ShowSeriesName;
            gdlo.ShowPercentage = this.ShowPercentage;
            gdlo.ShowBubbleSize = this.ShowBubbleSize;
            gdlo.Separator = this.Separator;
            gdlo.ShowLeaderLines = this.ShowLeaderLines;

            return gdlo;
        }
    }
}
