using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelTool.ZenrinIC.Models
{
    public class Highway
    {
        public int TempHighwayId { get; set; }
        public int PrefectureCode { get; set; }
        public string HighwayKana { get; set; }
        public string HighwayKanji { get; set; }
        public int InterchangeCount { get { return Interchanges.Count(); } }
        public List<InterchangeParsed> Interchanges { get; set; }
        public override string ToString()
        {
            return HighwayKanji;
        }
        private static Regex reGetOpenParentheses = new Regex(@"[\(|（].*$", RegexOptions.Compiled);
        private static Regex reGetOpenSpace= new Regex(@"\s.*$", RegexOptions.Compiled);
        public static IEnumerable<Highway> ParseHighways(IEnumerable<InterchangeParsed> interchanges)
        {
            // first pass removed 無料区間 and 均一区間
            var firstPass = interchanges
                .GroupBy(s => new
                {
                    PrefectureCode = s.PrefectureCode,
                    Highway = String.Join("／", s.HighwayKanji.Replace("／無料区間", "").Replace("（均一区間）", "").Split('／').Select(x => reGetOpenParentheses.Replace(x, "")).Distinct()),
                    HighwayKana = String.Join("/", s.HighwayKana.Replace(" ﾑﾘｮｳｸｶﾝ", "").Replace(" ｷﾝｲﾂｸｶﾝ", "").Split('/').Select(x => reGetOpenSpace.Replace(x, "")).Distinct()),
                })
                .Select(g => new Highway
                {
                    PrefectureCode = g.First().PrefectureCode,
                    HighwayKana = g.Key.HighwayKana,
                    HighwayKanji = g.Key.Highway,
                    Interchanges = g.ToList()
                })
                .ToList();
            var result = firstPass.Aggregate(new List<Highway>(), (acc, item) =>
            {
                var roads = item.HighwayKanji.Split('／');
                var roadKanas = item.HighwayKana.Split('/');
                for (var i = 0; i < roads.Length; i++)
                {
                    var existingHighway = acc.FirstOrDefault(s => s.HighwayKanji == roads[i]);
                    if (existingHighway == null)
                    {
                        existingHighway = new Highway
                        {
                            PrefectureCode = item.PrefectureCode,
                            HighwayKanji = roads[i],
                            HighwayKana = roadKanas[i],
                            Interchanges = item.Interchanges,
                        };
                        acc.Add(existingHighway);
                    }
                    else
                    {
                        existingHighway.Interchanges.AddRange(item.Interchanges);
                    }
                }
                return acc;
            });
            var index = 1;
            return result.OrderBy(s=>s.PrefectureCode).ThenBy(s=>s.HighwayKanji).Aggregate(new List<Highway>(), (acc, item)=>
            {
                item.TempHighwayId = index++;
                acc.Add(item);
                return acc;
            }).ToList();
        }
    }
}
