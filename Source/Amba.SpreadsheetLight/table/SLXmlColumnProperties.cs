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
using DocumentFormat.OpenXml.Spreadsheet;

namespace Amba.SpreadsheetLight
{
    internal class SLXmlColumnProperties
    {
        internal uint MapId { get; set; }
        internal string XPath { get; set; }
        internal bool? Denormalized { get; set; }
        internal XmlDataValues XmlDataType { get; set; }

        internal SLXmlColumnProperties()
        {
            this.SetAllNull();
        }

        private void SetAllNull()
        {
            this.MapId = 0;
            this.XPath = string.Empty;
            this.Denormalized = null;
            this.XmlDataType = XmlDataValues.AnyType;
        }

        internal void FromXmlColumnProperties(XmlColumnProperties xcp)
        {
            this.SetAllNull();

            if (xcp.MapId != null) this.MapId = xcp.MapId.Value;
            if (xcp.XPath != null) this.XPath = xcp.XPath.Value;
            if (xcp.Denormalized != null && xcp.Denormalized.Value) this.Denormalized = xcp.Denormalized.Value;
            if (xcp.XmlDataType != null) this.XmlDataType = xcp.XmlDataType.Value;
        }

        internal XmlColumnProperties ToXmlColumnProperties()
        {
            XmlColumnProperties xcp = new XmlColumnProperties();
            xcp.MapId = this.MapId;
            xcp.XPath = this.XPath;
            if (this.Denormalized != null && this.Denormalized.Value) xcp.Denormalized = this.Denormalized.Value;
            xcp.XmlDataType = this.XmlDataType;

            return xcp;
        }

        internal SLXmlColumnProperties Clone()
        {
            SLXmlColumnProperties xcp = new SLXmlColumnProperties();
            xcp.MapId = this.MapId;
            xcp.XPath = this.XPath;
            xcp.Denormalized = this.Denormalized;
            xcp.XmlDataType = this.XmlDataType;

            return xcp;
        }
    }
}
