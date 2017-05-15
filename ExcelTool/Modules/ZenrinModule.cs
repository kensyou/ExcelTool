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
using Mapsui.Layers;
using Mapsui.Styles;
using Mapsui.Providers;
using Mapsui.Geometries;
using Mapsui;
using Mapsui.Utilities;
using Mapsui.Projection;
using BruTile.Predefined;
using BruTile.Web;
using ExcelTool.UI.Model;
using System.IO;
using Microsoft.Win32;
using System.Windows;

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
                    SelectedPath = @"D:\dev\apac\tobedeleted\grid_jp_id\GRID.Schema.root\GRID_Upload.Schema\Utility\JpTransportInterchangeConverter\POI",
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

            var header = new List<object> { "PrefectureCode", "TempInterchangeId", "IC_Kana", "IC_Kanji", "Highway", "Latitude", "Longitude", "Data_Date" };
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

            var header = new List<object> { "HighwayId", "HighwayKanji", "HighwayKana", "InterchangeCount", "Interchanges" };
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

            var header = new List<object> { "HighwayId", "Highway", "Interchange", "SortOrder", "Latitude", "Longitude" };
            var data = new List<List<object>> { header }.Concat(highwayInterchanges.Select(r =>
                new List<object> { r.TempHighwayId, r.HighwayKanji, r.Interchange.IC_Kanji, r.SortOrder, r.Interchange.Latitude, r.Interchange.Longitude
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
        public IEnumerable<Highway> Highways
        {
            get
            {
                return _ZenrinParser.Highways;
            }
        }
        public async Task ExportInterchangeDataAsMergeSql()
        {
            if (_ZenrinParser.Highways == null || !_ZenrinParser.Highways.Any())
            {
                System.Windows.MessageBox.Show("Haven't loaded highways correctly", "Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            try
            {
                Microsoft.Win32.SaveFileDialog saveFileDialog1 = new Microsoft.Win32.SaveFileDialog();
                saveFileDialog1.Filter = "SQL files (*.sql)|*.sql";
                saveFileDialog1.RestoreDirectory = true;
                if (saveFileDialog1.ShowDialog() == true)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine(_ZenrinParser.MergeHighwayData(_ZenrinParser.Highways));
                    sb.AppendLine(_ZenrinParser.MergeInterchangeData(_ZenrinParser.Interchanges));

                    Stream stream;
                    if ((stream = saveFileDialog1.OpenFile()) != null)
                    {
                        var utf = Encoding.UTF8.GetBytes(sb.ToString());
                        stream.Write(utf, 0, utf.Length);
                        stream.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(string.Format("{0}\n{1}", "データ保存中にエラーが発生しました", ex.Message),
                    "データ保存エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private Highway SelectedHighway { get; set; }
        private MapStyle SelectedMapStyle { get; set; }
        public Func<Highway, MapStyle, Map> MapGenerator
        {
            get
            {
                return (Highway highway, MapStyle mapStyle) =>
                {
                    //TODO
                    if (highway != null) this.SelectedHighway = highway;
                    if (mapStyle != null) this.SelectedMapStyle = mapStyle;

                    var result = _ZenrinParser.HighwayInterchanges.Where(s => s.TempHighwayId == this.SelectedHighway.TempHighwayId).OrderBy(s => s.SortOrder).Select(ic =>
                  {
                      var label = ic.Interchange.IC_Kanji;
                      var center = new Mapsui.Geometries.Point(ic.Interchange.Longitude, ic.Interchange.Latitude);
                      // OSM uses spherical mercator coordinates. So transform the lon lat coordinates to spherical mercator
                      var sphericalMercatorCoordinate = SphericalMercator.FromLonLat(center.X, center.Y);
                      return Tuple.Create(label, sphericalMercatorCoordinate);
                  }).ToArray();

                    var map = new Map();
                    map.Layers.Add(Mapbox.CreateTileLayer(this.SelectedMapStyle));
                    // Set the center of the viewport to the coordinate. The UI will refresh automatically
                    //map.Viewport.Center = result[0].Item2;
                    // Additionally you might want to set the resolution, this could depend on your specific purpose
                    //map.Viewport.Resolution = 14;
                    var bb = new BoundingBox(result.Select(s => s.Item2).ToList());
                    //map.NavigateTo(bb, ScaleMethod.Fit);
                    ZoomHelper.ZoomToBoudingbox(map.Viewport, bb.Left, bb.Top, bb.Right, bb.Bottom, 640, 480, ScaleMethod.Fit);
                    //map.Viewport.Center = bb.GetCentroid();
                    //map.Viewport.Resolution = 13;
                    var points = result.Select(s => s.Item2).ToList();
                    map.Layers.Add(HighwayMap.CreateLineStringLayer(points, HighwayMap.CreateLineStringStyle()));
                    map.Layers.Add(HighwayMap.CreateHighwayICLayer(result));

                    return map;
                };
            }
        }
        public static class Mapbox
        {
            private static readonly BruTile.Attribution MapboxAttribution = new BruTile.Attribution(
                "© Mapbox", "http://www.mapbox.org/copyright");

            public static TileLayer CreateTileLayer(MapStyle mapStyle)
            {
                var mapId = mapStyle == null ? "mapbox.streets" : mapStyle.Name;
                var url = $"https://api.mapbox.com/v4/{mapId}/{{z}}/{{x}}/{{y}}@2x.png?access_token=pk.eyJ1Ijoia2Vuc3lvdSIsImEiOiJjajJvZjR3cDEwMmZ2MzNxYmNpMnZrc3FmIn0.h_3HKpIB-jFdH0Efi-rgkw";
                return new TileLayer(new HttpTileSource(new GlobalSphericalMercator(0, 18), url ,
                            new[] { "a", "b", "c" }, name: "Mapbox",
                            attribution: MapboxAttribution));
            }
        }
        public static class HighwayMap
        {
            public static ILayer CreateHighwayICLayer(Tuple<string, Mapsui.Geometries.Point>[] pts)
            {
                var memoryProvider = new MemoryProvider();
                foreach (var pt in pts)
                {
                    var featureWithDefaultStyle = new Feature { Geometry = pt.Item2 };
                    featureWithDefaultStyle.Styles.Add(new LabelStyle { Text = pt.Item1 });
                    memoryProvider.Features.Add(featureWithDefaultStyle);
                }
                return new MemoryLayer { Name = "高速IC", DataSource = memoryProvider };
            }


            public static ILayer CreateLineStringLayer(IEnumerable<Mapsui.Geometries.Point> pts, IStyle style = null)
            {
                return new MemoryLayer
                {
                    DataSource = new MemoryProvider(new Feature
                    {
                        Styles = new List<IStyle> { style },
                        Geometry = new LineString(pts)
                    }
                    ),
                    Name = "LineStringLayer",
                    Style = style
                };
            }

            public static IStyle CreateLineStringStyle()
            {
                return new VectorStyle
                {
                    Fill = null,
                    Outline = null,
                    Line = { Color = Color.Red, Width = 4 }
                };
            }
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
