using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Xl = NetOffice.ExcelApi;
using ExcelTool.Helper;
using ExcelTool.Forms;
using System.Windows.Forms;
using ExcelTool.ZenrinIC;
using ExcelTool.ZenrinIC.Models;
using NetOffice.ExcelApi.Enums;

namespace ExcelTool.Modules
{

    public class ZenrinModule : IZenrinModule
    {
        private readonly IExcelHelper _ExcelHelper;
        private ZenrinParser _ZenrinParser;
        public ZenrinModule(IExcelHelper excelHelper)
        {
            _ExcelHelper = excelHelper;
        }

        public async Task ImportInterchangeData()
        {
            try
            {
                _ExcelHelper.SwitchToBusyState();
                var folder = new FolderBrowserDialog()
                {
                    RootFolder = Environment.SpecialFolder.Desktop,
                    SelectedPath = @"D:\dev\apac\grid_jp_id\GRID.Schema.root\GRID_Upload.Schema\Utility\JpTransportInterchangeConverter\POI",
                    Description = "IC Source Folderを選択してください"
                };
                if (folder.ShowDialog() == DialogResult.OK)
                {
                    _ZenrinParser = new ZenrinParser();
                    _ZenrinParser.Load(folder.SelectedPath);
                    await _ExcelHelper.DownloadDataOnExcel("HighwayRaw", () => Task.FromResult(_ZenrinParser.Highways), WriteHighways);
                    await _ExcelHelper.DownloadDataOnExcel("InterchangeRaw", () => Task.FromResult(_ZenrinParser.Interchanges), WriteInterchanges);
                    await _ExcelHelper.DownloadDataOnExcel("HighwayInterchange", () => Task.FromResult(_ZenrinParser.HighwayInterchanges), WriteHighwayInterchanges);
                }
            }
            finally
            {
                _ExcelHelper.RestoreDefaultState();
            }
        }
        static string Format_YearMonth = "yyyy/MM";
        static string Format_DateTime = "yyyy/MM/dd HH:mm:ss";
        static string Format_Date = "yyyy/MM/dd";
        public Xl.Worksheet WriteInterchanges(Xl.Worksheet worksheet, IEnumerable<Interchange> interchanges)
        {
            var tableColOffset = 1;
            var tableRowOffset = 1;
            var titleCell = worksheet.Cells[1, 1];
            titleCell.Value = "Interchange";
            //titleCell.Style = _Style.GetMasterTableHeaderStyle(worksheet);
            titleCell.EntireColumn.ColumnWidth = 1.0;

            var header = new List<object> { "PrefectureCode", "TempInterchangeId", "IC_Kana", "IC_Kanji",  "Highway", "Latitude", "Longitude", "Data_Date" };
            var data = new List<List<object>> { header }.Concat(interchanges.Select(r =>
                new List<object> { r.PrefectureCode, r.TempInterchangeId, r.IC_Kana, r.IC_Kanji, r.HighwayDisplay, r.Latitude, r.Longitude
                , DateTime.SpecifyKind(r.DataDate, DateTimeKind.Utc).ToLocalTime()})).ToArray().CreateRectangularArray();
            var tableTopLeft = worksheet.Cells[tableRowOffset + 1, tableColOffset + 1];
            var tableBottomRight = worksheet.Cells[interchanges.Count() + tableRowOffset + 1, header.Count() + tableColOffset];
            Xl.Range range = worksheet.Range(tableTopLeft, tableBottomRight);
            range.set_Value(Type.Missing, data);
            var opList = worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, range, null, XlYesNoGuess.xlYes);
            opList.Name = $"Zenrin.Interchange";
            opList.TableStyle = "TableStyleLight9";

            opList.ListColumns[8].DataBodyRange.NumberFormatLocal = Format_DateTime;
            var colorRange = AddinContext.ExcelApp.Union(opList.ListColumns[1].DataBodyRange,
                opList.ListColumns[8].DataBodyRange);
            colorRange.Interior.ThemeColor = XlThemeColor.xlThemeColorAccent1;
            colorRange.Interior.TintAndShade = 0.5;

            opList.Range.Columns.AutoFit();
            opList.Range.Rows.AutoFit();
            return worksheet;
        }
        public Xl.Worksheet WriteHighways(Xl.Worksheet worksheet, IEnumerable<Highway> highways)
        {
            var tableColOffset = 1;
            var tableRowOffset = 1;
            var titleCell = worksheet.Cells[1, 1];
            titleCell.Value = "Highway";
            //titleCell.Style = _Style.GetMasterTableHeaderStyle(worksheet);
            titleCell.EntireColumn.ColumnWidth = 1.0;

            var header = new List<object> { "HighwayId", "HighwayKanji", "HighwayKana", "InterchangeCount", "Interchanges"};
            var data = new List<List<object>> { header }.Concat(highways.Select(r =>
                new List<object> { r.TempHighwayId, r.HighwayKanji, r.HighwayKana, r.InterchangeCount, String.Join(", ", r.Interchanges.Select(s=>s.IC_Kanji))
                })).ToArray().CreateRectangularArray();
            var tableTopLeft = worksheet.Cells[tableRowOffset + 1, tableColOffset + 1];
            var tableBottomRight = worksheet.Cells[highways.Count() + tableRowOffset + 1, header.Count() + tableColOffset];
            Xl.Range range = worksheet.Range(tableTopLeft, tableBottomRight);
            range.set_Value(Type.Missing, data);
            var opList = worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, range, null, XlYesNoGuess.xlYes);
            opList.Name = $"Zenrin.Highway";
            opList.TableStyle = "TableStyleLight9";

            var colorRange = opList.ListColumns[1].DataBodyRange;
            colorRange.Interior.ThemeColor = XlThemeColor.xlThemeColorAccent1;
            colorRange.Interior.TintAndShade = 0.5;

            opList.Range.Columns.AutoFit();
            opList.Range.Rows.AutoFit();
            return worksheet;
        }
        public Xl.Worksheet WriteHighwayInterchanges(Xl.Worksheet worksheet, IEnumerable<HighwayInterchange> highwayInterchanges)
        {
            var tableColOffset = 1;
            var tableRowOffset = 1;
            var titleCell = worksheet.Cells[1, 1];
            titleCell.Value = "HighwayInterchange";
            //titleCell.Style = _Style.GetMasterTableHeaderStyle(worksheet);
            titleCell.EntireColumn.ColumnWidth = 1.0;

            var header = new List<object> { "HighwayId", "Highway", "Interchange", "SortOrder", "Latitude", "Longitude"};
            var data = new List<List<object>> { header }.Concat(highwayInterchanges.Select(r =>
                new List<object> { r.TempHighwayId, r.HighwayKanji, r.Interchange.IC_Kanji, r.Interchange.SortOrder, r.Interchange.Latitude, r.Interchange.Longitude
                })).ToArray().CreateRectangularArray();
            var tableTopLeft = worksheet.Cells[tableRowOffset + 1, tableColOffset + 1];
            var tableBottomRight = worksheet.Cells[highwayInterchanges.Count() + tableRowOffset + 1, header.Count() + tableColOffset];
            Xl.Range range = worksheet.Range(tableTopLeft, tableBottomRight);
            range.set_Value(Type.Missing, data);
            var opList = worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, range, null, XlYesNoGuess.xlYes);
            opList.Name = $"Zenrin.HighwayInterchange";
            opList.TableStyle = "TableStyleLight9";

            var colorRange = opList.ListColumns[1].DataBodyRange;
            colorRange.Interior.ThemeColor = XlThemeColor.xlThemeColorAccent1;
            colorRange.Interior.TintAndShade = 0.5;

            opList.Range.Columns.AutoFit();
            opList.Range.Rows.AutoFit();
            return worksheet;
        }

    }
    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET function")]
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }
    }
}
