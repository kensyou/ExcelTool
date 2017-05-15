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
                    SelectedPath = @"N:\JP-TYO-IT_ProductandSolutions\Projects\GRID_ID\MasterData\1606DB\POI",
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
                    MergeHighwayData(sb, _ZenrinParser.Highways);
                    MergeInterchangeData(sb, _ZenrinParser.Interchanges);

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
        private void MergeHighwayData(StringBuilder sb, IEnumerable<Highway> highways)
        {
            sb.AppendLine(@"SET NOCOUNT ON
GO
CREATE TABLE #MAP_Highway 
(
     HighwayID         INT NOT NULL
    ,HighwayUID        UniqueIdentifier	NOT NULL DEFAULT NEWID()
	,StateID		INT NOT NULL 
	,SortOrder	INT NOT NULL
	,Highway_EN	NVARCHAR(250) NOT NULL
)
GO
CREATE TABLE #MAP_Highway_ML 
(
	HighwayID	INT		NOT NULL DEFAULT 1
	,LanguageID		INT					NOT NULL DEFAULT 1
	,Highway_ML	NVARCHAR(250)	NOT	NULL		-- Name or description
	,Kana	NVARCHAR(100)				-- kana
)
GO
CREATE TABLE #StateLookup 
(
    Id  INT NOT NULL
    ,State_EN   NVARCHAR(50) NOT NULL
    ,StateID    INT NOT NULL
)
GO
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(23, 'AICHI',227)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(5, 'AKITA',228)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(2, 'AOMORI',229)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(12, 'CHIBA',230)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(38, 'EHIME',231)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(18, 'FUKUI',232)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(40, 'FUKUOKA',233)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(7, 'FUKUSHIMA',234)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(21, 'GIFU',235)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(10, 'GUMMA',236)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(34, 'HIROSHIMA',237)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(1, 'HOKKAIDO',238)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(28, 'HYOGO',239)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(8, 'IBARAKI',240)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(17, 'ISHIKAWA',241)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(3, 'IWATE',242)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(37, 'KAGAWA',243)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(46, 'KAGOSHIMA',244)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(14, 'KANAGAWA',245)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(39, 'KOCHI',246)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(43, 'KUMAMOTO',247)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(26, 'KYOTO',248)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(24, 'MIE',249)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(4, 'MIYAGI',250)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(45, 'MIYAZAKI',251)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(20, 'NAGANO',252)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(42, 'NAGASAKI',253)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(29, 'NARA',254)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(15, 'NIIGATA',255)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(44, 'OITA',256)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(33, 'OKAYAMA',257)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(47, 'OKINAWA',258)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(27, 'OSAKA',259)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(41, 'SAGA',260)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(11, 'SAITAMA',261)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(25, 'SHIGA',262)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(32, 'SHIMANE',263)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(22, 'SHIZUOKA',264)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(9, 'TOCHIGI',265)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(36, 'TOKUSHIMA',266)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(13, 'TOKYO',267)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(31, 'TOTTORI',268)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(16, 'TOYAMA',269)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(30, 'WAKAYAMA',270)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(6, 'YAMAGATA',271)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(35, 'YAMAGUCHI',272)
INSERT INTO #StateLookup(Id, State_EN, StateID) VALUES(19, 'YAMANASHI',273)
GO

");
            var localIndex = 1;
            foreach (var r in highways.GroupBy(s => s.PrefectureCode))
            {
                sb.AppendLine($"DECLARE @StateID INT");
                sb.AppendLine($"SELECT @StateID = s.StateID FROM #StateLookup s WHERE Id = {r.Key}");
                foreach (var f in r)
                {
                    sb.AppendLine($"INSERT INTO #MAP_Highway (HighwayID, StateID, SortOrder, Highway_EN) VALUES ( {f.TempHighwayId}, @StateID, {localIndex++ * 10}, N'{f.HighwayKanji}')");
                    sb.AppendLine($"INSERT INTO #MAP_Highway_ML (HighwayID, LanguageID, Highway_ML, Kana) VALUES ( {f.TempHighwayId}, 100, N'{f.HighwayKanji}', N'{f.HighwayKana}')");
                }
                sb.AppendLine("GO");
                localIndex = 1;
            }
            sb.AppendLine("-- END OF HIGHWAY Generateion ");
            sb.AppendLine();
            sb.AppendLine();
        }
        private void MergeInterchangeData(StringBuilder sb, IEnumerable<Interchange> interchanges)
        {
            sb.AppendLine(@"
CREATE TABLE #MAP_TransportInterchange 
(
	 TransportInterchangeID 	INT NOT NULL
    ,TransportInterchangeUID	UniqueIdentifier	NOT NULL DEFAULT NEWID()
	,TransportInterchange_EN	NVARCHAR(50) NOT NULL
	,Location					GEOGRAPHY NOT NULL
)
GO
CREATE TABLE #MAP_TransportInterchange_ML 
(
	TransportInterchangeID		INT					NOT NULL 
	,LanguageID					INT					NOT NULL 
	,TransportInterchange_ML	NVARCHAR(250)		NOT NULL		-- Name or description
	,Kana						NVARCHAR(100) NULL
)
GO
CREATE TABLE #MAP_TransportInterchangeHighwayLNK 
(
    TransportInterchangeHighwayLNKUID  UniqueIdentifier	NOT NULL DEFAULT NEWID()
	,HighwayID 						INT NOT NULL		-- Set SEED>1 if predefined records
	,TransportInterchangeID  INT NOT NULL
	,SortOrder	INT NOT NULL
)
GO
");
            var localIndex = 1;
            foreach (var r in interchanges)
            {
                sb.AppendLine($"INSERT INTO #MAP_TransportInterchange (TransportInterchangeID, TransportInterchange_EN, Location) VALUES ( {r.TempInterchangeId},N'{r.IC_Kanji}', geography::Point({r.Latitude}, {r.Longitude}, 4326))");
                sb.AppendLine($"INSERT INTO #MAP_TransportInterchange_ML (TransportInterchangeID, LanguageID, TransportInterchange_ML, Kana) VALUES ( {r.TempInterchangeId}, 100, N'{r.IC_Kanji}', N'{r.IC_Kana}')");
                foreach (var highway in r.Highways)
                {
                    sb.AppendLine($"INSERT INTO #MAP_TransportInterchangeHighwayLNK (HighwayID, TransportInterchangeID, SortOrder) VALUES ( {highway.Item1.TempHighwayId}, {r.TempInterchangeId}, {highway.Item2 * 10})");
                }
                sb.AppendLine("GO");
                localIndex = 1;
            }
            sb.AppendLine("-- END OF Interchange Generateion ");
            sb.AppendLine();
            sb.AppendLine(@"
DISABLE TRIGGER [dbo].[MAP_TransportInterchangeHighwayLNK_Audit_TRG] ON [dbo].[MAP_TransportInterchangeHighwayLNK]
GO
DISABLE TRIGGER [dbo].[MAP_TransportInterchange_ML_Audit_TRG] ON [dbo].[MAP_TransportInterchange_ML]
GO
DISABLE TRIGGER [dbo].[MAP_TransportInterchange_Audit_TRG] ON [dbo].[MAP_TransportInterchange]
GO
DISABLE TRIGGER [dbo].[MAP_Highway_ML_Audit_TRG] ON [dbo].MAP_Highway_ML
GO
DISABLE TRIGGER [dbo].[MAP_Highway_Audit_TRG] ON [dbo].[MAP_Highway]
GO

-- MAP_Highway
MERGE [dbo].[MAP_Highway] AS target  
USING (SELECT HighwayID, HighwayUID, StateID, SortOrder, Highway_EN FROM #MAP_Highway AS sod  
    ) AS source (HighwayID, HighwayUID, StateID, SortOrder, Highway_EN)  
ON (target.StateID = source.StateID AND target.Highway_EN = source.Highway_EN)  
WHEN MATCHED   
    THEN UPDATE SET target.HighwayUID = source.HighwayUID,
					target.SortOrder = source.SortOrder,   
                    target.Modified = GETDATE(),
                    target.ModifiedBy = 'Bulk Update'
WHEN NOT MATCHED
	THEN INSERT (HighwayUID, StateID, SortOrder, Highway_EN) VALUES (source.HighwayUID, source.StateID, source.SortOrder, source.Highway_EN);
GO  

-- MAP_Highway_ML
MERGE [dbo].[MAP_Highway_ML] AS target  
USING (SELECT r.HighwayID, trm.LanguageID, trm.Highway_ML, trm.Kana 
FROM #MAP_Highway_ML trm
INNER JOIN #MAP_Highway tr ON tr.HighwayID = trm.HighwayID
INNER JOIN [dbo].[MAP_Highway] r ON r.HighwayUID = tr.HighwayUID
    ) AS source (HighwayID, LanguageID, Highway_ML, Kana )  
ON (target.HighwayID = source.HighwayID AND target.LanguageID = source.LanguageID AND target.Highway_ML = source.Highway_ML)  
WHEN MATCHED   
    THEN UPDATE SET target.Kana = source.Kana,   
                    target.Modified = GETDATE(),
                    target.ModifiedBy = 'Bulk Update'
WHEN NOT MATCHED
	THEN INSERT (HighwayID, LanguageID, Highway_ML, Kana) VALUES (source.HighwayID, source.LanguageID, source.Highway_ML, source.Kana);
GO  
-- MAP_TransportInterchange
MERGE [dbo].[MAP_TransportInterchange] AS target
USING (SELECT tempIC.TransportInterchangeUID, actualID.TransportInterchangeID, tempIC.TransportInterchange_EN, tempIC.Location FROM
	(SELECT  sod.TransportInterchangeUID, sod.TransportInterchange_EN, sod.Location FROM #MAP_TransportInterchange AS sod 
	INNER JOIN (
	SELECT  DISTINCT so.TransportInterchangeUID, so.TransportInterchange_EN FROM #MAP_TransportInterchange AS so 
	INNER JOIN (
			SELECT  ti.TransportInterchangeID FROM #MAP_TransportInterchange ti
			INNER JOIN #MAP_TransportInterchangeHighwayLNK lnk ON ti.TransportInterchangeID = lnk.TransportInterchangeID
			INNER JOIN #MAP_Highway r ON r.HighwayID = lnk.HighwayID
		) k ON so.TransportInterchangeID = k.TransportInterchangeID
	) kk ON sod.TransportInterchangeUID = kk.TransportInterchangeUID 
		) tempIC 
		LEFT JOIN (
		SELECT sod.TransportInterchangeID, sod.TransportInterchange_EN, sod.Location FROM dbo.MAP_TransportInterchange AS sod  
		INNER JOIN (
			SELECT ti.TransportInterchangeID FROM dbo.MAP_TransportInterchange ti
			INNER JOIN dbo.MAP_TransportInterchangeHighwayLNK lnk ON ti.TransportInterchangeID = lnk.TransportInterchangeID
			INNER JOIN dbo.MAP_Highway r ON r.HighwayID = lnk.HighwayID
		) k ON sod.TransportInterchangeID = k.TransportInterchangeID
		) actualID ON tempIC.TransportInterchange_EN = actualID.TransportInterchange_EN
) AS source (TransportInterchangeUID,TransportInterchangeID,TransportInterchange_EN, Location  )
ON (target.TransportInterchangeID = source.TransportInterchangeID)
WHEN MATCHED
	THEN UPDATE SET target.Location = source.Location,
					target.Modified = GETDATE(),
					target.ModifiedBy = 'Bulk Update'
WHEN NOT MATCHED
	THEN INSERT (CountryID, TransportInterchangeUID, TransportInterchange_EN, Location) VALUES (1035, source.TransportInterchangeUID, source.TransportInterchange_EN, source.Location);
GO

MERGE [dbo].[MAP_TransportInterchange_ML] AS target
USING (
	SELECT i.TransportInterchangeID, tim.LanguageID, tim.TransportInterchange_ML, tim.Kana 
	FROM #MAP_TransportInterchange_ML tim
	INNER JOIN #MAP_TransportInterchange ti ON tim.TransportInterchangeID = ti.TransportInterchangeID
	INNER JOIN [dbo].[MAP_TransportInterchange] i ON i.TransportInterchangeUID = ti.TransportInterchangeUID
) AS Source (TransportInterchangeID, LanguageID, TransportInterchange_ML, Kana)
ON (target.TransportInterchangeID = source.TransportInterchangeID AND target.LanguageID = source.LanguageID)  
WHEN MATCHED
	THEN UPDATE SET target.TransportInterchange_ML = source.TransportInterchange_ML,
					target.Kana = source.Kana,
					target.Modified = GETDATE(),
					target.ModifiedBy = 'Bulk Update'
WHEN NOT MATCHED
	THEN INSERT (TransportInterchangeID, LanguageID, TransportInterchange_ML, Kana) VALUES (source.TransportInterchangeID, source.LanguageID, source.TransportInterchange_ML, source.Kana);
GO

MERGE [dbo].[MAP_TransportInterchangeHighwayLNK] AS target
USING (
	SELECT DISTINCT r.HighwayID, i.TransportInterchangeID, tk.SortOrder
	FROM #MAP_TransportInterchangeHighwayLNK tk
	INNER JOIN #MAP_Highway tr ON tk.HighwayID = tr.HighwayID
	INNER JOIN #MAP_TransportInterchange ti ON tk.TransportInterchangeID = ti.TransportInterchangeID
	INNER JOIN [dbo].[MAP_Highway] r ON tr.HighwayUID = r.HighwayUID
	INNER JOIN [dbo].[MAP_TransportInterchange] i ON ti.TransportInterchangeUID = i.TransportInterchangeUID
) AS Source (HighwayID, TransportInterchangeID, SortOrder)
ON (target.HighwayID = source.HighwayID AND target.TransportInterchangeID = source.TransportInterchangeID)
WHEN MATCHED
	THEN UPDATE SET target.SortOrder = source.SortOrder,
					target.Modified = GETDATE(),
					target.ModifiedBy = 'Bulk Update'
WHEN NOT MATCHED
	THEN INSERT (HighwayID, TransportInterchangeID, SortOrder) VALUES (source.HighwayID, source.TransportInterchangeID, source.SortOrder);
GO

ENABLE TRIGGER [dbo].[MAP_TransportInterchangeHighwayLNK_Audit_TRG] ON [dbo].[MAP_TransportInterchangeHighwayLNK]
GO
ENABLE TRIGGER [dbo].[MAP_TransportInterchange_ML_Audit_TRG] ON [dbo].[MAP_TransportInterchange_ML]
GO
ENABLE TRIGGER [dbo].[MAP_TransportInterchange_Audit_TRG] ON [dbo].[MAP_TransportInterchange]
GO
ENABLE TRIGGER [dbo].[MAP_Highway_ML_Audit_TRG] ON [dbo].MAP_Highway_ML
GO
ENABLE TRIGGER [dbo].[MAP_Highway_Audit_TRG] ON [dbo].[MAP_Highway]
GO


");
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
