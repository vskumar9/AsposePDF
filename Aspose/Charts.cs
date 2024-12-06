using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Drawing;

namespace Aspose
{
    internal class Charts
    {
        public void CreateChart()
        {
            //Instantiating a Workbook object
            Workbook workbook = new Workbook();
            //Adding a new worksheet to the Excel object
            int sheetIndex = workbook.Worksheets.Add();
            //Obtaining the reference of the newly added worksheet by passing its sheet index
            Worksheet worksheet = workbook.Worksheets[sheetIndex];
            //Adding a sample value to "A1" cell
            worksheet.Cells["A1"].PutValue(50);
            //Adding a sample value to "A2" cell
            worksheet.Cells["A2"].PutValue(100);
            //Adding a sample value to "A3" cell
            worksheet.Cells["A3"].PutValue(150);
            //Adding a sample value to "A4" cell
            worksheet.Cells["A4"].PutValue(200);
            //Adding a sample value to "B1" cell
            worksheet.Cells["B1"].PutValue(60);
            //Adding a sample value to "B2" cell
            worksheet.Cells["B2"].PutValue(32);
            //Adding a sample value to "B3" cell
            worksheet.Cells["B3"].PutValue(50);
            //Adding a sample value to "B4" cell
            worksheet.Cells["B4"].PutValue(40);
            //Adding a sample value to "C1" cell as category data
            worksheet.Cells["C1"].PutValue("Q1");
            //Adding a sample value to "C2" cell as category data
            worksheet.Cells["C2"].PutValue("Q2");
            //Adding a sample value to "C3" cell as category data
            worksheet.Cells["C3"].PutValue("Y1");
            //Adding a sample value to "C4" cell as category data
            worksheet.Cells["C4"].PutValue("Y2");
            //Adding a chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            //Accessing the instance of the newly added chart
            Chart chart = worksheet.Charts[chartIndex];
            //Adding NSeries (chart data source) to the chart ranging from "A1" cell to "B4"
            chart.NSeries.Add("A1:B4", true);
            //Setting the data source for the category data of NSeries
            chart.NSeries.CategoryData = "C1:C4";
            //Setting the display unit of value(Y) axis.
            chart.ValueAxis.DisplayUnit = DisplayUnitType.Hundreds;
            DisplayUnitLabel displayUnitLabel = chart.ValueAxis.DisplayUnitLabel;
            //Setting the custom display unit label
            displayUnitLabel.Text = "100";
            //Saving the Excel file
            workbook.Save("book1.xls");
            //// Instantiating a Workbook object
            //Workbook workbook = new Workbook();

            //// Adding a new worksheet to the Excel object
            //int sheetIndex = workbook.Worksheets.Add();

            //// Obtaining the reference of the newly added worksheet by passing its sheet index
            //Worksheet worksheet = workbook.Worksheets[sheetIndex];

            //// Adding sample values to cells
            //worksheet.Cells["A1"].PutValue(50);
            //worksheet.Cells["A2"].PutValue(100);
            //worksheet.Cells["A3"].PutValue(150);
            //worksheet.Cells["A4"].PutValue(200);
            //worksheet.Cells["B1"].PutValue(60);
            //worksheet.Cells["B2"].PutValue(32);
            //worksheet.Cells["B3"].PutValue(50);
            //worksheet.Cells["B4"].PutValue(40);
            //worksheet.Cells["C1"].PutValue("Q1");
            //worksheet.Cells["C2"].PutValue("Q2");
            //worksheet.Cells["C3"].PutValue("Y1");
            //worksheet.Cells["C4"].PutValue("Y2");

            //// Adding a chart to the worksheet
            //int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);

            //// Accessing the instance of the newly added chart
            //Chart chart = worksheet.Charts[chartIndex];

            //// Adding NSeries (chart data source) to the chart ranging from "A1" cell to "B4"
            //int seriesIndex = chart.NSeries.Add("A1:B4", true);

            //// Setting the data source for the category data of NSeries
            //chart.NSeries.CategoryData = "C1:C4";

            //// Accessing the series and changing properties
            //Series series = chart.NSeries[seriesIndex];
            //series.Values = "=B1:B4"; // Setting the values of the series.

            //// Changing the chart type of the series
            //series.Type = ChartType.Line;

            //// Setting marker properties
            ////series.Marker.MarkerStyle = ChartMarkerType.Circle;
            ////series.Marker.ForegroundColor = Color.Black;
            ////series.Marker.BackgroundColor = Color.White;

            //series.Marker.MarkerStyle = ChartMarkerType.Circle;
            //series.Marker.ForegroundColorSetType = FormattingType.Automatic;
            //series.Marker.ForegroundColor = System.Drawing.Color.Black;
            //series.Marker.BackgroundColorSetType = FormattingType.Automatic;
            //series.Area.BackgroundColor = System.Drawing.Color.Black;


            //// Saving the Excel file
            //workbook.Save("book1.xls");
        }
    }
}
