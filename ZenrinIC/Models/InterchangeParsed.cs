using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.ZenrinIC.Models
{
    public class InterchangeParsed
    {
        public string FileName { get; set; }
        public int PrefectureCode { get; set; }
        public int RowNumber { get; set; }
        public int SortOrder { get; set; }
        public string ZenrinTypeCode { get; set; }
        public int PrefectureICSerial { get; set; }
        public string HighwayKana { get; set; }
        public string IC_Kana { get; set; }
        public string HighwayKanji { get; set; }
        public string IC_Kanji { get; set; }
        public double Latitude { get; set; }
        public double Longitude { get; set; }
        public DateTime DataDate { get; set; }
    }
}
