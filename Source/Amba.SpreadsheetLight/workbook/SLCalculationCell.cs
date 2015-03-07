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
    internal class SLCalculationCell
    {
        internal int RowIndex { get; set; }
        internal int ColumnIndex { get; set; }
        internal int SheetId { get; set; }
        internal bool? InChildChain { get; set; }
        internal bool? NewLevel { get; set; }
        internal bool? NewThread { get; set; }
        internal bool? Array { get; set; }

        internal SLCalculationCell()
        {
            this.SetAllNull();
        }

        internal SLCalculationCell(string CellReference)
        {
            this.SetAllNull();

            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                this.RowIndex = iRowIndex;
                this.ColumnIndex = iColumnIndex;
            }
        }

        private void SetAllNull()
        {
            this.RowIndex = 1;
            this.ColumnIndex = 1;
            this.SheetId = 0;
            this.InChildChain = null;
            this.NewLevel = null;
            this.NewThread = null;
            this.Array = null;
        }

        internal void FromCalculationCell(CalculationCell cc)
        {
            this.SetAllNull();

            int iRowIndex = -1;
            int iColumnIndex = -1;
            if (SLTool.FormatCellReferenceToRowColumnIndex(cc.CellReference.Value, out iRowIndex, out iColumnIndex))
            {
                this.RowIndex = iRowIndex;
                this.ColumnIndex = iColumnIndex;
            }


            this.SheetId = cc.SheetId ?? 0;
            if (cc.InChildChain != null) this.InChildChain = cc.InChildChain.Value;
            if (cc.NewLevel != null) this.NewLevel = cc.NewLevel.Value;
            if (cc.NewThread != null) this.NewThread = cc.NewThread.Value;
            if (cc.Array != null) this.Array = cc.Array.Value;
        }

        internal CalculationCell ToCalculationCell()
        {
            CalculationCell cc = new CalculationCell();
            cc.CellReference = SLTool.ToCellReference(this.RowIndex, this.ColumnIndex);
            cc.SheetId = this.SheetId;
            if (this.InChildChain != null && this.InChildChain.Value) cc.InChildChain = this.InChildChain.Value;
            if (this.NewLevel != null && this.NewLevel.Value) cc.NewLevel = this.NewLevel.Value;
            if (this.NewThread != null && this.NewThread.Value) cc.NewThread = this.NewThread.Value;
            if (this.Array != null && this.Array.Value) cc.Array = this.Array.Value;

            return cc;
        }

        internal SLCalculationCell Clone()
        {
            SLCalculationCell cc = new SLCalculationCell();
            cc.RowIndex = this.RowIndex;
            cc.ColumnIndex = this.ColumnIndex;
            cc.SheetId = this.SheetId;
            cc.InChildChain = this.InChildChain;
            cc.NewLevel = this.NewLevel;
            cc.NewThread = this.NewThread;
            cc.Array = this.Array;

            return cc;
        }
    }
}
