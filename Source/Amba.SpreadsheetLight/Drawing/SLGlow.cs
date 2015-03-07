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
using SLA = Amba.SpreadsheetLight.Drawing;

namespace Amba.SpreadsheetLight.Drawing
{
    /// <summary>
    /// Encapsulates properties and methods for specifying glow effects.
    /// This simulates the DocumentFormat.OpenXml.Drawing.Glow class.
    /// </summary>
    public class SLGlow
    {
        internal bool HasGlow;

        private decimal decRadius;
        internal decimal Radius
        {
            get { return decRadius; }
            set
            {
                decRadius = value;
                if (decRadius < 0m) decRadius = 0m;
                if (decRadius > 2147483647m) decRadius = 2147483647m;
            }
        }
        internal SLColorTransform GlowColor { get; set; }

        internal SLGlow(List<System.Drawing.Color> ThemeColors)
        {
            this.HasGlow = false;
            this.decRadius = 0;
            this.GlowColor = new SLColorTransform(ThemeColors);
        }

        /// <summary>
        /// Set no glow.
        /// </summary>
        public void SetNoGlow()
        {
            this.HasGlow = false;
            this.decRadius = 0;
        }

        /// <summary>
        /// Set the glow color.
        /// </summary>
        /// <param name="GlowColor">The color used.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Size">The size in points. The suggested range is 0 pt to 150 pt (both inclusive).</param>
        public void SetGlow(System.Drawing.Color GlowColor, decimal Transparency, decimal Size)
        {
            this.HasGlow = true;
            this.GlowColor.SetColor(GlowColor, Transparency);
            this.Radius = Size;
        }

        /// <summary>
        /// Set the glow color.
        /// </summary>
        /// <param name="GlowColor">The theme color used.</param>
        /// <param name="Tint">The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color and positive tints lighten the theme color.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Size">The size in points. The suggested range is 0 pt to 150 pt (both inclusive).</param>
        public void SetGlow(SLThemeColorIndexValues GlowColor, double Tint, decimal Transparency, decimal Size)
        {
            this.HasGlow = true;
            this.GlowColor.SetColor(GlowColor, Tint, Transparency);
            this.Radius = Size;
        }

        internal A.Glow ToGlow()
        {
            A.Glow g = new A.Glow();
            if (this.GlowColor.IsRgbColorModelHex)
            {
                g.RgbColorModelHex = this.GlowColor.ToRgbColorModelHex();
            }
            else
            {
                g.SchemeColor = this.GlowColor.ToSchemeColor();
            }
            g.Radius = SLA.SLDrawingTool.CalculatePositiveCoordinate(this.decRadius);

            return g;
        }

        internal SLGlow Clone()
        {
            SLGlow g = new SLGlow(this.GlowColor.listThemeColors);
            g.HasGlow = this.HasGlow;
            g.decRadius = this.decRadius;
            g.GlowColor = this.GlowColor.Clone();

            return g;
        }
    }
}
