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
    /// <summary>
    /// Encapsulates properties and methods for representing a merged cell range. This simulates the DocumentFormat.OpenXml.Spreadsheet.MergeCell class.
    /// The actual merging of cells is done by a SLDocument function. This class is for supporting purposes.
    /// </summary>
    public class SLMergeCell
    {
        internal int iStartRowIndex;
        /// <summary>
        /// The row index of the top row in the merged cell range. This is read-only.
        /// </summary>
        public int StartRowIndex
        {
            get { return iStartRowIndex; }
        }

        internal int iStartColumnIndex;
        /// <summary>
        /// The column index of the left column in the merged cell range. This is read-only.
        /// </summary>
        public int StartColumnIndex
        {
            get { return iStartColumnIndex; }
        }

        internal int iEndRowIndex;
        /// <summary>
        /// The row index of the bottom row in the merged cell range. This is read-only.
        /// </summary>
        public int EndRowIndex
        {
            get { return iEndRowIndex; }
        }

        internal int iEndColumnIndex;
        /// <summary>
        /// The column index of the right column in the merged cell range. This is read-only.
        /// </summary>
        public int EndColumnIndex
        {
            get { return iEndColumnIndex; }
        }

        private bool bIsValid;
        /// <summary>
        /// Indicates if the merged cell range is valid. This is read-only.
        /// </summary>
        public bool IsValid
        {
            get { return bIsValid; }
        }

        /// <summary>
        /// Initializes an instance of SLMergeCell.
        /// </summary>
        public SLMergeCell()
        {
            iStartRowIndex = 1;
            iStartColumnIndex = 1;
            iEndRowIndex = 1;
            iEndColumnIndex = 1;
            bIsValid = false;
        }

        /// <summary>
        /// Form a SLMergeCell given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the top-right corner cell.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the corner cell.</param>
        /// <param name="StartColumnIndex">The column index of the corner cell.</param>
        /// <param name="EndRowIndex">The row index of the opposite corner cell.</param>
        /// <param name="EndColumnIndex">The column index of the opposite corner cell.</param>
        public void FromIndices(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            if (iStartRowIndex == iEndRowIndex && iStartColumnIndex == iEndColumnIndex)
            {
                // it's the same cell! We'll treat this as invalid.
                this.bIsValid = false;
            }
            else
            {
                this.bIsValid = SLTool.CheckRowColumnIndexLimit(iStartRowIndex, iStartColumnIndex) && SLTool.CheckRowColumnIndexLimit(iEndRowIndex, iEndColumnIndex);
            }
        }

        /// <summary>
        /// Form a SLMergeCell from a DocumentFormat.OpenXml.Spreadsheet.MergeCell class.
        /// </summary>
        /// <param name="mc">The source DocumentFormat.OpenXml.Spreadsheet.MergeCell class.</param>
        public void FromMergeCell(MergeCell mc)
        {
            string sStartCell = string.Empty, sEndCell = string.Empty;
            int index = 0;
            bool bStartSuccess = false, bEndSuccess = false;
            bIsValid = false;

            if (mc.Reference != null)
            {
                index = mc.Reference.Value.IndexOf(":");
                // if "A1:C3", then the index must be at least at the 3rd position (or index 2)
                if (index >= 2)
                {
                    sStartCell = mc.Reference.Value.Substring(0, index);
                    sEndCell = mc.Reference.Value.Substring(index + 1);

                    bStartSuccess = SLTool.FormatCellReferenceToRowColumnIndex(sStartCell, out this.iStartRowIndex, out this.iStartColumnIndex);
                    bEndSuccess = SLTool.FormatCellReferenceToRowColumnIndex(sEndCell, out this.iEndRowIndex, out this.iEndColumnIndex);

                    if (bStartSuccess && bEndSuccess)
                    {
                        bIsValid = true;
                    }
                }
            }
        }

        /// <summary>
        /// Form a DocumentFormat.OpenXml.Spreadsheet.MergeCell class from this SLMergeCell class.
        /// </summary>
        /// <returns>A DocumentFormat.OpenXml.Spreadsheet.MergeCell class.</returns>
        public MergeCell ToMergeCell()
        {
            MergeCell mc = new MergeCell();
            string sStartCell = SLTool.ToCellReference(iStartRowIndex, iStartColumnIndex);
            string sEndCell = SLTool.ToCellReference(iEndRowIndex, iEndColumnIndex);
            mc.Reference = string.Format("{0}:{1}", sStartCell, sEndCell);

            return mc;
        }

        internal SLMergeCell Clone()
        {
            SLMergeCell mc = new SLMergeCell();
            mc.iStartRowIndex = this.iStartRowIndex;
            mc.iStartColumnIndex = this.iStartColumnIndex;
            mc.iEndRowIndex = this.iEndRowIndex;
            mc.iEndColumnIndex = this.iEndColumnIndex;
            mc.bIsValid = this.bIsValid;

            return mc;
        }
    }
}
