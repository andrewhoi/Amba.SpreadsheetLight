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

namespace Amba.SpreadsheetLight.Drawing
{
    internal class SLGradientStop
    {
        internal SLColorTransform Color { get; set; }
        private decimal decPosition;
        /// <summary>
        /// The position in percentage ranging from 0% to 100%. Accurate to 1/1000 of a percent.
        /// </summary>
        internal decimal Position
        {
            get { return decPosition; }
            set
            {
                decPosition = value;
                if (decPosition < 0m) decPosition = 0m;
                if (decPosition > 100m) decPosition = 100m;
            }
        }

        internal SLGradientStop(List<System.Drawing.Color> ThemeColors)
        {
            this.Color = new SLColorTransform(ThemeColors);
            this.Position = 0m;
        }

        internal SLGradientStop(List<System.Drawing.Color> ThemeColors, string HexColor, decimal Position)
        {
            this.Color = new SLColorTransform(ThemeColors);
            this.Position = Position;

            System.Drawing.Color clr = new System.Drawing.Color();
            try
            {
                clr = System.Drawing.Color.FromArgb(int.Parse(HexColor, System.Globalization.NumberStyles.HexNumber));
            }
            catch
            {
                clr = System.Drawing.Color.White;
            }
            this.Color.SetColor(clr, 0);
        }

        internal A.GradientStop ToGradientStop()
        {
            A.GradientStop gs = new A.GradientStop();
            if (this.Color.IsRgbColorModelHex) gs.RgbColorModelHex = this.Color.ToRgbColorModelHex();
            else gs.SchemeColor = this.Color.ToSchemeColor();

            gs.Position = Convert.ToInt32(this.Position * 1000m);

            return gs;
        }

        internal SLGradientStop Clone()
        {
            SLGradientStop gs = new SLGradientStop(this.Color.listThemeColors);
            gs.Color = this.Color.Clone();
            gs.decPosition = this.decPosition;

            return gs;
        }
    }
}
