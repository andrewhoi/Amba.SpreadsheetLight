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
    internal class SLSheet
    {
        internal string Name { get; set; }
        internal uint SheetId { get; set; }
        internal SheetStateValues State { get; set; }
        internal string Id { get; set; }
        internal SLSheetType SheetType { get; set; }

        internal SLSheet(string Name, uint SheetId, string Id, SLSheetType SheetType)
        {
            this.Name = Name;
            this.SheetId = SheetId;
            this.State = SheetStateValues.Visible;
            this.Id = Id;
            this.SheetType = SheetType;
        }

        internal Sheet ToSheet()
        {
            Sheet s = new Sheet();
            s.Name = this.Name;
            s.SheetId = this.SheetId;
            if (this.State != SheetStateValues.Visible) s.State = this.State;
            s.Id = this.Id;

            return s;
        }

        internal SLSheet Clone()
        {
            SLSheet s = new SLSheet(this.Name, this.SheetId, this.Id, this.SheetType);
            s.State = this.State;
            return s;
        }
    }
}
