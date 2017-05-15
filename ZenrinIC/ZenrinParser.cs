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
            HighwayInterchanges = Interchanges.SelectMany(ic=>ic.Highways.Select(h=> new { Highway = h.Item1, SortOrder = h.Item2, Interchange = ic })).OrderBy(hi=>hi.Highway.TempHighwayId).ThenBy(hi=>hi.SortOrder)
                .Select(i => new HighwayInterchange { Interchange = i.Interchange, SortOrder = i.SortOrder, HighwayKana = i.Highway.HighwayKana, PrefectureCode = i.Highway.PrefectureCode, HighwayKanji = i.Highway.HighwayKanji, TempHighwayId = i.Highway.TempHighwayId }).ToList();
            return this;
        }
    }
}
