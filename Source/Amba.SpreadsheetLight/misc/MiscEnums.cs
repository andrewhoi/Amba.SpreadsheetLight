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

namespace Amba.SpreadsheetLight
{
    /// <summary>
    /// The type of hyperlink.
    /// </summary>
    public enum SLHyperlinkTypeValues
    {
        /// <summary>
        /// Hyperlink to an existing web page.
        /// </summary>
        Url = 0,
        /// <summary>
        /// Hyperlink to an existing file.
        /// </summary>
        FilePath,
        /// <summary>
        /// Hyperlink to a place within the spreadsheet (cell references or defined names).
        /// </summary>
        InternalDocumentLink,
        /// <summary>
        /// Hyperlink to an email address.
        /// </summary>
        EmailAddress
    }

    /// <summary>
    /// The type of measurement unit.
    /// </summary>
    public enum SLMeasureUnitTypeValues
    {
        /// <summary>
        /// English Metric Unit. No, not the bird...
        /// </summary>
        Emu = 0,
        /// <summary>
        /// Inch.
        /// </summary>
        Inch,
        /// <summary>
        /// Centimeter.
        /// </summary>
        Centimeter,
        /// <summary>
        /// Point.
        /// </summary>
        Point
    }

    /// <summary>
    /// The type of paste options.
    /// </summary>
    public enum SLPasteTypeValues
    {
        /// <summary>
        /// Just plain pasting. Fanfare and choral singing each sold separately. *smile*
        /// </summary>
        Paste = 0,
        /// <summary>
        /// Paste only values (no formulas).
        /// </summary>
        Values,
        /// <summary>
        /// Paste values and formulas. NOTE: Formulas are copied as is (no recalculating cell references).
        /// </summary>
        Formulas,
        /// <summary>
        /// Transpose.
        /// </summary>
        Transpose,
        /// <summary>
        /// Paste only formatting (styles).
        /// </summary>
        Formatting
    }
}
