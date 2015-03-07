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
    /// Encapsulates properties and methods for specifying effects such as glow, shadows, reflection and soft edges.
    /// This simulates the DocumentFormat.OpenXml.Drawing.EffectList class.
    /// </summary>
    public class SLEffectList
    {
        internal List<System.Drawing.Color> listThemeColors;

        internal bool HasEffectList
        {
            get
            {
                return this.Glow.HasGlow || this.Shadow.IsInnerShadow != null
                    || this.Reflection.HasReflection || this.SoftEdge.HasSoftEdge;
            }
        }

        // A.Blur is not accessible from Excel! Don't know what values to allow...

        internal SLGlow Glow { get; set; }

        internal SLShadowEffect Shadow { get; set; }

        internal SLReflection Reflection { get; set; }

        internal SLSoftEdge SoftEdge { get; set; }

        internal SLEffectList(List<System.Drawing.Color> ThemeColors)
        {
            int i;
            this.listThemeColors = new List<System.Drawing.Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
            {
                this.listThemeColors.Add(ThemeColors[i]);
            }

            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.Glow = new SLGlow(this.listThemeColors);
            this.Shadow = new SLShadowEffect(this.listThemeColors);
            this.Reflection = new SLReflection();
            this.SoftEdge = new SLSoftEdge();
        }

        internal A.EffectList ToEffectList()
        {
            A.EffectList el = new A.EffectList();

            if (this.Glow.HasGlow)
            {
                el.Glow = this.Glow.ToGlow();
            }

            if (this.Shadow.IsInnerShadow != null)
            {
                if (this.Shadow.IsInnerShadow.Value)
                {
                    el.InnerShadow = this.Shadow.ToInnerShadow();
                }
                else
                {
                    el.OuterShadow = this.Shadow.ToOuterShadow();
                }
            }

            if (this.Reflection.HasReflection)
            {
                el.Reflection = this.Reflection.ToReflection();
            }

            if (this.SoftEdge.HasSoftEdge)
            {
                el.SoftEdge = this.SoftEdge.ToSoftEdge();
            }

            return el;
        }

        internal SLEffectList Clone()
        {
            SLEffectList el = new SLEffectList(this.listThemeColors);
            el.Glow = this.Glow.Clone();
            el.Shadow = this.Shadow.Clone();
            el.Reflection = this.Reflection.Clone();
            el.SoftEdge = this.SoftEdge.Clone();

            return el;
        }
    }
}
