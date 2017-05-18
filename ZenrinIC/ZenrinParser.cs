using ExcelTool.ZenrinIC.Models;
using LumenWorks.Framework.IO.Csv;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.ZenrinIC
{
    public class ZenrinParser
    {
        private DirectoryInfo _Folder;
        public IEnumerable<Highway> Highways { get; set; }
        public IEnumerable<Interchange> Interchanges { get; set; }
        public IEnumerable<HighwayInterchange> HighwayInterchanges { get; set; }

        public ZenrinParser()
        {
            SqlServerTypes.Utilities.LoadNativeAssemblies(AppDomain.CurrentDomain.BaseDirectory);
        }
        public ZenrinParser Load(string selectedFolder)
        {
            Highways = new List<Highway>();
            Interchanges = new List<Interchange>();
            HighwayInterchanges = new List<HighwayInterchange>();
            _Folder = new DirectoryInfo(selectedFolder);
            var targetFiles = _Folder.GetFiles("*.TXT", SearchOption.TopDirectoryOnly).ToList();
            var data = new List<InterchangeParsed>();
            foreach (var tf in targetFiles)
            {
                using (FileStream fs = File.OpenRead(tf.FullName))
                using (TextReader reader = new StreamReader(fs, Encoding.GetEncoding("shift_jis")))
                {
                    using (var rodadCsv = new CachedCsvReader(reader, false, '\t'))
                    {
                        data.AddRange(rodadCsv.Select((s, i) =>
                        {
                            return new InterchangeRaw
                            {
                                FileName = tf.Name,
                                RowNumber = i + 1,
                                PrefectureCode = s.GetValue(0).ToString(),
                                ZenrinTypeCode = s.GetValue(1).ToString(),
                                PrefectureICSerial = s.GetValue(2).ToString(),
                                HighwayKana = s.GetValue(6).ToString(),
                                IC_Kana = s.GetValue(7).ToString(),
                                HighwayKanji = s.GetValue(8).ToString(),
                                IC_Kanji = s.GetValue(9).ToString(),
                                Longitude = s.GetValue(17).ToString(),
                                Latitude = s.GetValue(18).ToString(),
                                DataDate = s.GetValue(30).ToString(),
                            }.Parse();
                        }).ToList());
                    }
                }
            }
            Highways = Highway.ParseHighways(data);
            Interchanges = Interchange.NormalizedInterchange(Highways);
            HighwayInterchanges = Interchanges.SelectMany(ic => ic.Highways.Select(h => new { Highway = h.Item1, SortOrder = h.Item2, Interchange = ic })).OrderBy(hi => hi.Highway.TempHighwayId).ThenBy(hi => hi.SortOrder)
                .Select(i => new HighwayInterchange { Interchange = i.Interchange, SortOrder = i.SortOrder, HighwayKana = i.Highway.HighwayKana, PrefectureCode = i.Highway.PrefectureCode, HighwayKanji = i.Highway.HighwayKanji, TempHighwayId = i.Highway.TempHighwayId }).ToList();
            return this;
        }

        private string GetTempStateLookup()
        {
            return @"CREATE TABLE #StateLookup 
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
GO";
        }

        public string MergeHighwayData(IEnumerable<Highway> highways)
        {
            var sb = new StringBuilder();
            sb.AppendLine(@"SET NOCOUNT ON
GO
CREATE TABLE #MAP_Highway 
(
     HighwayID         INT NOT NULL
    ,HighwayUID        UniqueIdentifier	NOT NULL DEFAULT NEWID()
	,CountryID		INT NOT NULL 
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
");
            var localIndex = 1;
            foreach (var f in highways)
            {
                sb.AppendLine($"INSERT INTO #MAP_Highway (HighwayID, CountryID, SortOrder, Highway_EN) VALUES ( {f.TempHighwayId}, 1035, {localIndex++ * 10}, N'{f.HighwayKanji}')");
                sb.AppendLine($"INSERT INTO #MAP_Highway_ML (HighwayID, LanguageID, Highway_ML, Kana) VALUES ( {f.TempHighwayId}, 100, N'{f.HighwayKanji}', N'{f.HighwayKana}')");
            }
            sb.AppendLine("GO");
            localIndex = 1;
            sb.AppendLine("-- END OF HIGHWAY Generateion ");
            sb.AppendLine();
            sb.AppendLine();
            return sb.ToString();
        }

        public string MergeInterchangeData(IEnumerable<Interchange> interchanges)
        {
            var sb = new StringBuilder();
            sb.AppendLine(@"
CREATE TABLE #MAP_TransportInterchange 
(
	 TransportInterchangeID 	INT NOT NULL
    ,TransportInterchangeUID	UniqueIdentifier	NOT NULL DEFAULT NEWID()
	,TransportInterchange_EN	NVARCHAR(50) NOT NULL
    ,StateTempID                    INT NOT NULL
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
            sb.AppendLine(GetTempStateLookup());

            var localIndex = 1;
            foreach (var r in interchanges)
            {
                sb.AppendLine($"INSERT INTO #MAP_TransportInterchange (TransportInterchangeID, TransportInterchange_EN, StateTempID, Location) VALUES ( {r.TempInterchangeId},N'{r.IC_Kanji}', {r.PrefectureCode}, geography::Point({r.Latitude}, {r.Longitude}, 4326))");
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
USING (SELECT HighwayID, HighwayUID, CountryID, SortOrder, Highway_EN FROM #MAP_Highway AS sod  
    ) AS source (HighwayID, HighwayUID, CountryID, SortOrder, Highway_EN)  
ON (target.CountryID = source.CountryID AND target.Highway_EN = source.Highway_EN)  
WHEN MATCHED   
    THEN UPDATE SET target.HighwayUID = source.HighwayUID,
					target.SortOrder = source.SortOrder,   
                    target.Modified = GETDATE(),
                    target.ModifiedBy = 'Bulk Update'
WHEN NOT MATCHED
	THEN INSERT (HighwayUID, CountryID, SortOrder, Highway_EN) VALUES (source.HighwayUID, source.CountryID, source.SortOrder, source.Highway_EN);
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
USING (
	SELECT sl.StateID, tempIC.TransportInterchangeUID, actualID.TransportInterchangeID, tempIC.TransportInterchange_EN, tempIC.Location 
	FROM
	(SELECT sod.StateTempID, sod.TransportInterchangeUID, sod.TransportInterchange_EN, sod.Location FROM #MAP_TransportInterchange AS sod 
	INNER JOIN (
	SELECT  DISTINCT so.TransportInterchangeUID, so.TransportInterchange_EN FROM #MAP_TransportInterchange AS so 
	INNER JOIN (
			SELECT  ti.TransportInterchangeID FROM #MAP_TransportInterchange ti
			INNER JOIN #MAP_TransportInterchangeHighwayLNK lnk ON ti.TransportInterchangeID = lnk.TransportInterchangeID
			INNER JOIN #MAP_Highway r ON r.HighwayID = lnk.HighwayID
		) k ON so.TransportInterchangeID = k.TransportInterchangeID
	) kk ON sod.TransportInterchangeUID = kk.TransportInterchangeUID 
		) tempIC 
	INNER JOIN #StateLookup sl ON tempIC.StateTempID = sl.Id
	LEFT JOIN dbo.MAP_TransportInterchange actualID ON sl.StateID = actualID.StateID AND tempIC.TransportInterchange_EN = actualID.TransportInterchange_EN
) AS source (StateID, TransportInterchangeUID,TransportInterchangeID,TransportInterchange_EN, Location)
ON (target.TransportInterchangeID = source.TransportInterchangeID)
WHEN MATCHED
	THEN UPDATE SET target.Location = source.Location,
					target.StateID = source.StateID,
					target.TransportInterchangeUID = source.TransportInterchangeUID,
					target.Modified = GETDATE(),
					target.ModifiedBy = 'Bulk Update'
WHEN NOT MATCHED
	THEN INSERT (StateID, TransportInterchangeUID, TransportInterchange_EN, Location) VALUES (source.StateID, source.TransportInterchangeUID, source.TransportInterchange_EN, source.Location);
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
            return sb.ToString();
        }

        public string SqlExtractForCsv()
        {
            var sb = new StringBuilder();
            sb.AppendLine(GetTempStateLookup());
            sb.Append(@"SELECT lnk.TransportInterchangeHighwayLNKID AS id, ti.TransportInterchange_EN AS title, r.Highway_EN AS title_road, ti.TransportInterchangeID AS ic_key, sl.Id AS prefecture_id, r.HighwayID road_id, lnk.SortOrder AS list_order, ti.Location.Lat As lat_wgs84, ti.Location.Long AS lon_wgs84
FROM MAP_TransportInterchangeHighwayLNK lnk
INNER JOIN MAP_Highway r ON lnk.HighwayID = r.HighwayID
INNER JOIN MAP_TransportInterchange ti on lnk.TransportInterchangeID = ti.TransportInterchangeID
INNER JOIN #StateLookup sl ON sl.StateID = ti.StateID
ORDER BY r.HighwayID, lnk.SortOrder,ti.StateID");
            return sb.ToString();
        }
    }
}
