using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelTool.ZenrinIC.Models
{
    public class HighwayInterchange
    {
        public int TempHighwayId { get; set; }
        public int PrefectureCode { get; set; }
        public string HighwayKana { get; set; }
        public string HighwayKanji { get; set; }
        public InterchangeParsed Interchange { get; set; }
        public override string ToString()
        {
            return HighwayKanji + " - " + Interchange.IC_Kanji;
        }
        
    }
}
