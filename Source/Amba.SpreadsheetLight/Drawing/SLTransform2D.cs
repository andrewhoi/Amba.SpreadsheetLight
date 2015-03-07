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
using A = DocumentFormat.OpenXml.Drawing;

namespace Amba.SpreadsheetLight.Drawing
{
    internal class SLTransform2D
    {
        internal bool HasOffset;
        internal SLOffset Offset { get; set; }
        internal bool HasExtents;
        internal SLExtents Extents { get; set; }

        internal int? Rotation { get; set; }
        internal bool? HorizontalFlip { get; set; }
        internal bool? VerticalFlip { get; set; }

        internal SLTransform2D()
        {
            this.HasOffset = false;
            this.Offset = new SLOffset();
            this.HasExtents = false;
            this.Extents = new SLExtents();

            this.Rotation = null;
            this.HorizontalFlip = null;
            this.VerticalFlip = null;
        }

        internal A.Transform2D ToTransform2D()
        {
            A.Transform2D trans = new A.Transform2D();
            if (this.HasOffset) trans.Offset = this.Offset.ToOffset();
            if (this.HasExtents) trans.Extents = this.Extents.ToExtents();

            if (this.Rotation != null) trans.Rotation = this.Rotation.Value;
            if (this.HorizontalFlip != null) trans.HorizontalFlip = this.HorizontalFlip.Value;
            if (this.VerticalFlip != null) trans.VerticalFlip = this.VerticalFlip.Value;

            return trans;
        }

        internal SLTransform2D Clone()
        {
            SLTransform2D trans = new SLTransform2D();
            trans.HasOffset = this.HasOffset;
            trans.Offset = this.Offset.Clone();
            trans.HasExtents = this.HasExtents;
            trans.Extents = this.Extents.Clone();

            trans.Rotation = this.Rotation;
            trans.HorizontalFlip = this.HorizontalFlip;
            trans.VerticalFlip = this.VerticalFlip;

            return trans;
        }
    }
}
