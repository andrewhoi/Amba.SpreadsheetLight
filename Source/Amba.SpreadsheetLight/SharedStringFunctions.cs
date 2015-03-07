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
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Amba.SpreadsheetLight
{
    public partial class SLDocument
    {
        /// <summary>
        /// Get existing shared strings. WARNING: This is only a snapshot. Any changes made to the returned result are not used.
        /// </summary>
        /// <returns>A list of existing shared strings.</returns>
        public List<SLRstType> GetSharedStrings()
        {
            List<SLRstType> result = new List<SLRstType>();
            SLRstType rst = new SLRstType();
            for (int i = 0; i < listSharedString.Count; ++i)
            {
                rst.FromHash(listSharedString[i]);
                result.Add(rst.Clone());
            }

            return result;
        }

        /// <summary>
        /// Get existing shared strings in a list of SharedStringItem objects. WARNING: This is only a snapshot. Any changes made to the returned result are not used.
        /// </summary>
        /// <returns>A list of existing SharedStringItem objects.</returns>
        public List<SharedStringItem> GetSharedStringItems()
        {
            List<SharedStringItem> result = new List<SharedStringItem>();
            SharedStringItem ssi;

            for (int i = 0; i < listSharedString.Count; ++i)
            {
                ssi = new SharedStringItem();
                ssi.InnerXml = listSharedString[i];
                result.Add(ssi);
            }

            return result;
        }

        internal void LoadSharedStringTable()
        {
            countSharedString = 0;
            listSharedString = new List<string>();
            dictSharedStringHash = new Dictionary<string, int>();

            if (wbp.SharedStringTablePart != null)
            {
                OpenXmlReader oxr = OpenXmlReader.Create(wbp.SharedStringTablePart);
                while (oxr.Read())
                {
                    if (oxr.ElementType == typeof(SharedStringItem))
                    {
                        this.ForceSaveToSharedStringTable((SharedStringItem)oxr.LoadCurrentElement());
                    }
                }
                oxr.Dispose();

                countSharedString = listSharedString.Count;
            }
        }

        internal void WriteSharedStringTable()
        {
            if (wbp.SharedStringTablePart != null)
            {
                if (listSharedString.Count > countSharedString)
                {
                    wbp.SharedStringTablePart.SharedStringTable.Count = (uint)listSharedString.Count;
                    wbp.SharedStringTablePart.SharedStringTable.UniqueCount = (uint)dictSharedStringHash.Count;

                    int diff = listSharedString.Count - countSharedString;
                    for (int i = 0; i < diff; ++i)
                    {
                        wbp.SharedStringTablePart.SharedStringTable.Append(new SharedStringItem()
                        {
                            InnerXml = listSharedString[i + countSharedString]
                        });
                    }

                    wbp.SharedStringTablePart.SharedStringTable.Save();
                }
            }
            else
            {
                if (listSharedString.Count > 0)
                {
                    SharedStringTablePart sstp = wbp.AddNewPart<SharedStringTablePart>();
                    using (MemoryStream ms = new MemoryStream())
                    {
                        using (StreamWriter sw = new StreamWriter(ms))
                        {
                            sw.Write("<x:sst count=\"{0}\" uniqueCount=\"{1}\" xmlns:x=\"{2}\">", listSharedString.Count, dictSharedStringHash.Count, SLConstants.NamespaceX);
                            for (int i = 0; i < listSharedString.Count; ++i)
                            {
                                sw.Write("<x:si>{0}</x:si>", listSharedString[i]);
                            }
                            sw.Write("</x:sst>");
                            sw.Flush();
                            ms.Position = 0;
                            sstp.FeedData(ms);
                        }
                    }
                }
            }
        }
    }
}
