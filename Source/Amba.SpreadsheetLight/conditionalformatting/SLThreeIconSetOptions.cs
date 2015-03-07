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

namespace Amba.SpreadsheetLight
{
    /// <summary>
    /// Conditional formatting options for three icon sets.
    /// </summary>
    public class SLThreeIconSetOptions
    {
        internal SLThreeIconSetValues IconSetType { get; set; }

        /// <summary>
        /// Specifies if the icons in the set are reversed.
        /// </summary>
        public bool ReverseIconOrder { get; set; }

        /// <summary>
        /// Specifies if only the icon is shown. Set to false to show both icon and value.
        /// </summary>
        public bool ShowIconOnly { get; set; }

        internal bool IsCustomIcon;

        internal SLIconValues vIcon1;
        /// <summary>
        /// The 1st icon.
        /// </summary>
        public SLIconValues Icon1
        {
            get { return vIcon1; }
            set
            {
                if (vIcon1 != value)
                {
                    vIcon1 = value;
                    IsCustomIcon = true;
                }
            }
        }

        internal SLIconValues vIcon2;
        /// <summary>
        /// The 2nd icon.
        /// </summary>
        public SLIconValues Icon2
        {
            get { return vIcon2; }
            set
            {
                if (vIcon2 != value)
                {
                    vIcon2 = value;
                    IsCustomIcon = true;
                }
            }
        }

        internal SLIconValues vIcon3;
        /// <summary>
        /// The 3rd icon.
        /// </summary>
        public SLIconValues Icon3
        {
            get { return vIcon3; }
            set
            {
                if (vIcon3 != value)
                {
                    vIcon3 = value;
                    IsCustomIcon = true;
                }
            }
        }

        /// <summary>
        /// Specifies if values are to be greater than or equal to the 2nd range value. Set to false if values are to be strictly greater than.
        /// </summary>
        public bool GreaterThanOrEqual2 { get; set; }

        /// <summary>
        /// Specifies if values are to be greater than or equal to the 3rd range value. Set to false if values are to be strictly greater than.
        /// </summary>
        public bool GreaterThanOrEqual3 { get; set; }

        /// <summary>
        /// The 2nd range value.
        /// </summary>
        public string Value2 { get; set; }

        /// <summary>
        /// The 3rd range value.
        /// </summary>
        public string Value3 { get; set; }

        /// <summary>
        /// The conditional format type for the 2nd range value.
        /// </summary>
        public SLConditionalFormatRangeValues Type2 { get; set; }

        /// <summary>
        /// The conditional format type for the 3rd range value.
        /// </summary>
        public SLConditionalFormatRangeValues Type3 { get; set; }

        /// <summary>
        /// Initializes an instance of SLThreeIconSetOptions.
        /// </summary>
        /// <param name="IconSetType">The type of icon set.</param>
        public SLThreeIconSetOptions(SLThreeIconSetValues IconSetType)
        {
            this.IconSetType = IconSetType;
            this.ReverseIconOrder = false;
            this.ShowIconOnly = false;

            this.IsCustomIcon = false;

            switch (IconSetType)
            {
                case SLThreeIconSetValues.ThreeArrows:
                    this.vIcon1 = SLIconValues.RedDownArrow;
                    this.vIcon2 = SLIconValues.YellowSideArrow;
                    this.vIcon3 = SLIconValues.GreenUpArrow;
                    break;
                case SLThreeIconSetValues.ThreeArrowsGray:
                    this.vIcon1 = SLIconValues.GrayDownArrow;
                    this.vIcon2 = SLIconValues.GraySideArrow;
                    this.vIcon3 = SLIconValues.GrayUpArrow;
                    break;
                case SLThreeIconSetValues.ThreeFlags:
                    this.vIcon1 = SLIconValues.RedFlag;
                    this.vIcon2 = SLIconValues.YellowFlag;
                    this.vIcon3 = SLIconValues.GreenFlag;
                    break;
                case SLThreeIconSetValues.ThreeSigns:
                    this.vIcon1 = SLIconValues.RedDiamond;
                    this.vIcon2 = SLIconValues.YellowTriangle;
                    this.vIcon3 = SLIconValues.GreenCircle;
                    break;
                case SLThreeIconSetValues.ThreeStars:
                    this.vIcon1 = SLIconValues.SilverStar;
                    this.vIcon2 = SLIconValues.HalfGoldStar;
                    this.vIcon3 = SLIconValues.GoldStar;
                    break;
                case SLThreeIconSetValues.ThreeSymbols:
                    this.vIcon1 = SLIconValues.RedCrossSymbol;
                    this.vIcon2 = SLIconValues.YellowExclamationSymbol;
                    this.vIcon3 = SLIconValues.GreenCheckSymbol;
                    break;
                case SLThreeIconSetValues.ThreeSymbols2:
                    this.vIcon1 = SLIconValues.RedCross;
                    this.vIcon2 = SLIconValues.YellowExclamation;
                    this.vIcon3 = SLIconValues.GreenCheck;
                    break;
                case SLThreeIconSetValues.ThreeTrafficLights1:
                    this.vIcon1 = SLIconValues.RedCircleWithBorder;
                    this.vIcon2 = SLIconValues.YellowCircle;
                    this.vIcon3 = SLIconValues.GreenCircle;
                    break;
                case SLThreeIconSetValues.ThreeTrafficLights2:
                    this.vIcon1 = SLIconValues.RedTrafficLight;
                    this.vIcon2 = SLIconValues.YellowTrafficLight;
                    this.vIcon3 = SLIconValues.GreenTrafficLight;
                    break;
                case SLThreeIconSetValues.ThreeTriangles:
                    this.vIcon1 = SLIconValues.RedDownTriangle;
                    this.vIcon2 = SLIconValues.YellowDash;
                    this.vIcon3 = SLIconValues.GreenUpTriangle;
                    break;
            }

            this.GreaterThanOrEqual2 = true;
            this.GreaterThanOrEqual3 = true;

            this.Value2 = "33";
            this.Value3 = "67";

            this.Type2 = SLConditionalFormatRangeValues.Percent;
            this.Type3 = SLConditionalFormatRangeValues.Percent;
        }
    }
}
