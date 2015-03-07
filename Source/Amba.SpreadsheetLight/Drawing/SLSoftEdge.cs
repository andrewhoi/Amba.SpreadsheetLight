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
using SLA = Amba.SpreadsheetLight.Drawing;

namespace Amba.SpreadsheetLight.Drawing
{
    /// <summary>
    /// Encapsulates properties and methods for specifying soft edges.
    /// This simulates the DocumentFormat.OpenXml.Drawing.SoftEdge class.
    /// </summary>
    public class SLSoftEdge
    {
        internal bool HasSoftEdge;

        internal decimal decRadius;
        // Probably an example of the marketing team, usability/testing team, technical documentation
        // team and the development team not talking to each other properly.
        // "Normal" people relate to the word "Size". Technical people have no problems with "Radius".
        // I'm gonna go with the technical documentation slash development team here.
        // Also, the Open XML specs use radius. Ahh... but the tech doc people wrote the specs...
        /// <summary>
        /// Also known as "Size", and is measured in points. The suggested range is 0 pt to 100 pt (both inclusive).
        /// </summary>
        public decimal Radius
        {
            get { return decRadius; }
            set
            {
                HasSoftEdge = true;
                decRadius = value;
                if (decRadius < 0m) decRadius = 0m;
                if (decRadius > 2147483647m) decRadius = 2147483647m;
            }
        }

        /// <summary>
        /// Set no soft edge.
        /// </summary>
        public void SetNoSoftEdge()
        {
            this.HasSoftEdge = false;
            this.decRadius = 0;
        }

        internal SLSoftEdge()
        {
            this.HasSoftEdge = false;
            this.decRadius = 0;
        }

        internal A.SoftEdge ToSoftEdge()
        {
            A.SoftEdge se = new A.SoftEdge();
            se.Radius = SLA.SLDrawingTool.CalculatePositiveCoordinate(decRadius);

            return se;
        }

        internal SLSoftEdge Clone()
        {
            SLSoftEdge se = new SLSoftEdge();
            se.HasSoftEdge = this.HasSoftEdge;
            se.decRadius = this.decRadius;

            return se;
        }
    }
}
