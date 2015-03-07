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
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Amba.SpreadsheetLight.Charts
{
    /// <summary>
    /// Chart customization options for bubble charts.
    /// </summary>
    public class SLBubbleChartOptions
    {
        /// <summary>
        /// Specifies if the bubbles have a 3D effect.
        /// </summary>
        public bool Bubble3D { get; set; }

        internal uint iBubbleScale;
        /// <summary>
        /// Scale factor in percentage of the default size, ranging from 0% to 300% (both inclusive). The default is 100%.
        /// </summary>
        public uint BubbleScale
        {
            get { return iBubbleScale; }
            set
            {
                iBubbleScale = value;
                if (iBubbleScale > 300) iBubbleScale = 300;
            }
        }

        /// <summary>
        /// Specifies if negatively sized bubbles are shown.
        /// </summary>
        public bool ShowNegativeBubbles { get; set; }

        /// <summary>
        /// Specifies how bubble sizes relate to the presentation of the bubbles.
        /// </summary>
        public C.SizeRepresentsValues SizeRepresents { get; set; }

        /// <summary>
        /// Initializes an instance of SLBubbleChartOptions.
        /// </summary>
        public SLBubbleChartOptions()
        {
            this.Bubble3D = true;
            this.iBubbleScale = 100;
            this.ShowNegativeBubbles = true;
            this.SizeRepresents = C.SizeRepresentsValues.Area;
        }

        internal SLBubbleChartOptions Clone()
        {
            SLBubbleChartOptions bco = new SLBubbleChartOptions();
            bco.Bubble3D = this.Bubble3D;
            bco.iBubbleScale = this.iBubbleScale;
            bco.ShowNegativeBubbles = this.ShowNegativeBubbles;
            bco.SizeRepresents = this.SizeRepresents;

            return bco;
        }
    }
}
