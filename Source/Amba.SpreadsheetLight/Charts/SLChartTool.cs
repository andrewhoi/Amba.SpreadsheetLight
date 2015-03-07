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
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Amba.SpreadsheetLight.Charts
{
    internal class SLChartTool
    {
        internal static bool IsSuitableCategoryHeader(Dictionary<SLCellPoint, SLCell> Cells, int RowIndex, int ColumnIndex)
        {
            bool result = false;
            SLCellPoint pt = new SLCellPoint(RowIndex, ColumnIndex);
            if (Cells.ContainsKey(pt))
            {
                if (Cells[pt].DataType == CellValues.String || Cells[pt].DataType == CellValues.SharedString)
                {
                    result = true;
                }
            }

            return result;
        }

        internal static string GetChartReferenceFormula(string WorksheetName, int RowIndex, int ColumnIndex)
        {
            return SLTool.ToCellReference(WorksheetName, RowIndex, ColumnIndex, true);
        }

        internal static string GetChartReferenceFormula(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            return SLTool.ToCellRange(WorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, true);
        }

        internal static bool IsCombinationChartFriendly(SLDataSeriesChartType ChartType)
        {
            bool result = true;

            switch (ChartType)
            {
                case SLDataSeriesChartType.Area3DChart:
                case SLDataSeriesChartType.Bar3DChart:
                case SLDataSeriesChartType.BubbleChart:
                case SLDataSeriesChartType.Line3DChart:
                case SLDataSeriesChartType.Pie3DChart:
                case SLDataSeriesChartType.SurfaceChart:
                case SLDataSeriesChartType.Surface3DChart:
                case SLDataSeriesChartType.StockChart:
                case SLDataSeriesChartType.None:
                    result = false;
                    break;
            }

            return result;
        }

        internal static bool Is3DChart(SLAreaChartType ChartType)
        {
            bool result = false;
            switch (ChartType)
            {
                case SLAreaChartType.Area3D:
                case SLAreaChartType.StackedArea3D:
                case SLAreaChartType.StackedAreaMax3D:
                    result = true;
                    break;
            }

            return result;
        }

        internal static bool Is3DChart(SLBarChartType ChartType)
        {
            bool result = false;
            switch (ChartType)
            {
                case SLBarChartType.ClusteredBar3D:
                case SLBarChartType.ClusteredHorizontalCone:
                case SLBarChartType.ClusteredHorizontalCylinder:
                case SLBarChartType.ClusteredHorizontalPyramid:
                case SLBarChartType.StackedBar3D:
                case SLBarChartType.StackedBarMax3D:
                case SLBarChartType.StackedHorizontalCone:
                case SLBarChartType.StackedHorizontalConeMax:
                case SLBarChartType.StackedHorizontalCylinder:
                case SLBarChartType.StackedHorizontalCylinderMax:
                case SLBarChartType.StackedHorizontalPyramid:
                case SLBarChartType.StackedHorizontalPyramidMax:
                    result = true;
                    break;
            }

            return result;
        }

        internal static bool Is3DChart(SLBubbleChartType ChartType)
        {
            // all bubble charts are 2D
            return false;
        }

        internal static bool Is3DChart(SLColumnChartType ChartType)
        {
            bool result = false;
            switch (ChartType)
            {
                case SLColumnChartType.ClusteredColumn3D:
                case SLColumnChartType.ClusteredCone:
                case SLColumnChartType.ClusteredCylinder:
                case SLColumnChartType.ClusteredPyramid:
                case SLColumnChartType.Column3D:
                case SLColumnChartType.Cone3D:
                case SLColumnChartType.Cylinder3D:
                case SLColumnChartType.Pyramid3D:
                case SLColumnChartType.StackedColumn3D:
                case SLColumnChartType.StackedColumnMax3D:
                case SLColumnChartType.StackedCone:
                case SLColumnChartType.StackedConeMax:
                case SLColumnChartType.StackedCylinder:
                case SLColumnChartType.StackedCylinderMax:
                case SLColumnChartType.StackedPyramid:
                case SLColumnChartType.StackedPyramidMax:
                    result = true;
                    break;
            }

            return result;
        }

        internal static bool Is3DChart(SLDoughnutChartType ChartType)
        {
            // all doughnut charts are 2D
            return false;
        }

        internal static bool Is3DChart(SLLineChartType ChartType)
        {
            bool result = false;
            switch (ChartType)
            {
                case SLLineChartType.Line3D:
                    result = true;
                    break;
            }

            return result;
        }

        internal static bool Is3DChart(SLPieChartType ChartType)
        {
            // while there are 3D pie versions, there are no floors, sidewalls, backwalls.
            return false;
        }

        internal static bool Is3DChart(SLRadarChartType ChartType)
        {
            // all radar charts are 2D
            return false;
        }

        internal static bool Is3DChart(SLScatterChartType ChartType)
        {
            // all scatter charts are 2D
            return false;
        }

        internal static bool Is3DChart(SLStockChartType ChartType)
        {
            // all stock charts are 2D
            return false;
        }

        internal static bool Is3DChart(SLSurfaceChartType ChartType)
        {
            // all surface charts are 3D
            // However, the contour charts don't show the side and back walls, only the floor.
            // You know, because it's in orthogonal view.
            return true;
        }

        internal static C.AxisPositionValues GetOppositePosition(C.AxisPositionValues Position)
        {
            C.AxisPositionValues result = Position;
            switch (Position)
            {
                case C.AxisPositionValues.Bottom:
                    result = C.AxisPositionValues.Top;
                    break;
                case C.AxisPositionValues.Left:
                    result = C.AxisPositionValues.Right;
                    break;
                case C.AxisPositionValues.Right:
                    result = C.AxisPositionValues.Left;
                    break;
                case C.AxisPositionValues.Top:
                    result = C.AxisPositionValues.Bottom;
                    break;
            }

            return result;
        }
    }
}
