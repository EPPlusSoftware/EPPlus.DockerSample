using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System.Dynamic;
using System.Text.Json;
namespace EPPlus.DockerSamples.Alpine
{
    public static class ExcelReport
    {
        public static byte[] GetReport()
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            using (var p = new ExcelPackage())
            {

                CreateCoverPage(p);
                
                var data = GetData();
                CreateTableAndChartWorksheet(p, data);

                CreatePivotTable(p);

                return p.GetAsByteArray();
            }
        }

        private static void CreateCoverPage(ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("Cover Page");
            //var pic = ws.Drawings.AddPicture("Logo", "resources/logo.png");
            //pic.SetPosition(0, 5, 0, 5);
            ws.View.ShowGridLines = false;
            ws.View.PageLayoutView = true;
            ws.View.ZoomScale = 100;
            ws.HeaderFooter.OddHeader.InsertPicture(new FileInfo("resources/logo.png"), PictureAlignment.Centered);
            var textbox = ws.Drawings.AddShape("Text1", eShapeStyle.Rect);
            textbox.SetPosition(5, 0, 0, 2);
            textbox.SetSize(500, 250);
            textbox.TextAnchoring = eTextAnchoringType.Top;
            textbox.TextAlignment = eTextAlignment.Center;
            
            var rt1=textbox.RichText.Add("EPPlus Docker Sample\r\n",true);
            rt1.LatinFont = "Verdana";
            rt1.Size = 24;
            var rt2 = textbox.RichText.Add("This sample generates a report showing different features of EPPlus in a variety of environments. EPPlus will work seamlessly in all environments, even when no access to GDI is available, like in Windows Nano Server, Linux without libgdiplus and even in web assembly inside the browser. This worksheet is in page layout view to show the logo picture added in the header.");
            rt2.LatinFont = "Verdana";
            rt2.Size = 16;

            //Add a named style for the hyperlink;
            var ns=p.Workbook.Styles.CreateNamedStyle("Hyperlink");
            ns.BuildInId = 8; //Build in type 8 is Hyperlink
            ns.Style.Font.Color.SetColor(eThemeSchemeColor.Hyperlink);
            ns.Style.Font.UnderLine = true;


            var hyperlink = new ExcelHyperLink("https://samples.epplussoftware.com/BlazorSamples");
            hyperlink.Display = "See Blazor samples here...";

            ws.Cells["D20"].Hyperlink = hyperlink;
            ws.Cells["D20"].StyleName = "Hyperlink";
            ws.Columns[4].Width = 20;

            var pic = ws.Drawings.AddPicture("EPPlusHQ","resources/office.jpg");
            pic.SetPosition(25, 1, 1, 5);
            pic.SetSize(200);


            ws.Cells["A24"].Value = "Sample of how to add an image: EPPlus HQ located in WTC, Stockholm";
            ws.Cells["A24"].Style.Font.Size = 15;


            var nextWorksheetHyperlink = new ExcelHyperLink("'FX Rates'!A1", "Go to next worksheet...");
            ws.Cells["D44"].Hyperlink = nextWorksheetHyperlink;
            ws.Cells["D44"].StyleName = "Hyperlink";
            
            ws.Columns[4].Width = 20;
        }

        private static IEnumerable<ExpandoObject>? GetData()
        {
            string json = File.ReadAllText("resources/FxRates.json");
            if (json == null) throw new InvalidDataException("json can't be null");
            var data = JsonSerializer.Deserialize<IEnumerable<ExpandoObject>>(json);
            return data;
        }

        private static void CreateTableAndChartWorksheet(ExcelPackage p,IEnumerable<ExpandoObject>? data)
        {
            var ws = p.Workbook.Worksheets.Add("FX Rates");
            ws.View.ShowGridLines = false;

            var range = ws.Cells["A20"].LoadFromDictionaries(data, x =>
            {
                x.TableStyle = null;
                x.DataTypes = new[] { eDataTypes.DateTime, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number };
                x.PrintHeaders = true;
            });

            ws.Cells["A:A"].Style.Numberformat.Format = "yyyy-MM-dd";
            ws.Cells["B:F"].Style.Numberformat.Format = "#,##0.00";

            ws.Cells[22, 7, range.End.Row, 11].Formula = "B$21/B22-1";
            ws.Cells["G:K"].Style.Numberformat.Format = "0.00%";

            ws.Cells["B20"].Value = "USD/SEK";
            ws.Cells["C20"].Value = "USD/EUR";
            ws.Cells["D20"].Value = "USD/INR";
            ws.Cells["E20"].Value = "USD/CNY";
            ws.Cells["F20"].Value = "USD/DKK";

            ws.Cells["G20"].Value = "USD/SEK %";
            ws.Cells["H20"].Value = "USD/EUR %";
            ws.Cells["I20"].Value = "USD/INR %";
            ws.Cells["J20"].Value = "USD/CNY %";
            ws.Cells["K20"].Value = "USD/DKK %";

            //Add a table over the range including the . 
            var tbl = ws.Tables.Add(ws.Cells[20, 1, range.End.Row, 11], "Table1");
            tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Dark6;

            ws.View.FreezePanes(21, 1);

            CreateLineChart(ws, range);

            var nextWorksheetHyperlink = new ExcelHyperLink("'Pivot table'!A6", "Go to next worksheet...");
            ws.Cells["M19"].Hyperlink = nextWorksheetHyperlink;
            ws.Cells["M19"].StyleName = "Hyperlink";

            ws.Cells.AutoFitColumns();
        }

        private static void CreateLineChart(ExcelWorksheet ws, ExcelRangeBase range)
        {
            var chart = ws.Drawings.AddLineChart("LineChart1", eLineChartType.Line);
            var s1 = chart.Series.Add(ws.Cells[21, 7, range.End.Row, 7], ws.Cells[21, 1, range.End.Row, 1]);
            s1.HeaderAddress = ws.Cells["G20"];

            var s2 = chart.Series.Add(ws.Cells[21, 8, range.End.Row, 8], ws.Cells[21, 1, range.End.Row, 1]);
            s2.HeaderAddress = ws.Cells["H20"];

            var s3 = chart.Series.Add(ws.Cells[21, 9, range.End.Row, 9], ws.Cells[21, 1, range.End.Row, 1]);
            s3.HeaderAddress = ws.Cells["I20"];

            var s4 = chart.Series.Add(ws.Cells[21, 10, range.End.Row, 10], ws.Cells[21, 1, range.End.Row, 1]);
            s4.HeaderAddress = ws.Cells["J20"];

            var s5 = chart.Series.Add(ws.Cells[21, 11, range.End.Row, 11], ws.Cells[21, 1, range.End.Row, 1]);
            s5.HeaderAddress = ws.Cells["K20"];

            chart.XAxis.Crosses = eCrosses.Min;

            chart.To.Row = 19;
            chart.To.Column = 11;
            chart.Legend.Add();
            chart.Legend.Position = eLegendPosition.Bottom;
            chart.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.LineChartStyle9);
        }

        private static void CreatePivotTable(ExcelPackage p)
        {
            var wsPivot = p.Workbook.Worksheets.Add("Pivot table");
            var wsFxRates = p.Workbook.Worksheets["Fx Rates"];
            wsPivot.Cells["A1"].Value = "This pivot table uses the six first columns of the table in the FX Rates worksheet as source. Data is shown as monthly average over year 2017.";
            var pt = wsPivot.PivotTables.Add(wsPivot.Cells["A3"], wsFxRates.Cells[20,1,wsFxRates.Dimension.End.Row,6], "PivotTable1");

            var rf = pt.ColumnFields.Add(pt.Fields[0]);
            rf.AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months, new DateTime(2016, 6, 1), new DateTime(2017, 12, 1));
            for(int i=1;i<=5;i++)
            {
                var df = pt.DataFields.Add(pt.Fields[i]);
                df.Function = DataFieldFunctions.Average;
                df.Format = "0.000";
            }
            pt.ColumnHeaderCaption = "Average Monthly Rates";
            pt.GrandTotalCaption = "Average Rates";
            pt.PivotTableStyle = OfficeOpenXml.Table.PivotTableStyles.Medium13;

            var shape = wsPivot.Drawings.AddShape("SampleLink", eShapeStyle.Rect);            
            
            shape.SetPosition(11, 0, 1, 5);
            shape.SetSize(200, 50);
            shape.Hyperlink=new Uri("https://github.com/EPPlusSoftware/EPPlus.Sample.NetCore");
            shape.TextAlignment = eTextAlignment.Center;
            shape.Font.UnderLine = eUnderLineType.Single;
            shape.Text = "For more sample on how to use EPPlus, please checkout our sample project at Github...";

        }
    }
}
