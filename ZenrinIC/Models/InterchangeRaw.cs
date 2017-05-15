using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.ZenrinIC.Models
{
    public class InterchangeRaw
    {
        public string FileName { get; set; }
        public int RowNumber { get; set; }
        public string PrefectureCode { get; set; }
        public string ZenrinTypeCode { get; set; }
        public string PrefectureICSerial { get; set; }
        public string HighwayKana { get; set; }
        public string IC_Kana { get; set; }
        public string HighwayKanji { get; set; }
        public string IC_Kanji { get; set; }
        public string Latitude { get; set; }
        public string Longitude { get; set; }
        public string DataDate { get; set; }

        public InterchangeParsed Parse()
        {
            try
            {
                var r = new InterchangeParsed();
                r.FileName = this.FileName;
                r.RowNumber = this.RowNumber;
                r.PrefectureCode = Int32.Parse(this.PrefectureCode.TrimStart('0'));
                r.ZenrinTypeCode = this.ZenrinTypeCode;
                r.PrefectureICSerial = Int32.Parse(this.PrefectureICSerial.TrimStart('0'));
                r.HighwayKana = this.HighwayKana;
                r.IC_Kana = this.IC_Kana;
                r.HighwayKanji = this.HighwayKanji;
                r.IC_Kanji = this.IC_Kanji;
                var lat_t = ConvertDegreeAngleToDouble(this.Latitude);
                var lon_t = ConvertDegreeAngleToDouble(this.Longitude);
                r.Latitude = ConvertJapToWgs84Lat(lat_t, lon_t);
                r.Longitude = ConvertJapToWgs84Long(lat_t, lon_t);
                r.DataDate = DateTime.ParseExact(this.DataDate, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.AssumeLocal);
                return r;
            }
            catch (Exception ex)
            {
                throw new Exception($"FileName: {FileName}, Row: {RowNumber}, Exception: {ex.Message}");
            }
        }
        private static double ConvertJapToWgs84Lat(double lat_t, double lon_t)
        {
            return lat_t - lat_t * 0.00010695 + lon_t * 0.000017464 + 0.0046017;
        }
        private static double ConvertJapToWgs84Long(double lat_t, double lon_t)
        {
            return lon_t - lat_t * 0.000046038 - lon_t * 0.000083043 + 0.010040;
        }
        public static double ConvertDegreeAngleToDouble(string degressMinutesSeconds)
        {
            var d = degressMinutesSeconds.Split(':').Select(s => Double.Parse(s)).ToArray();
            return ConvertDegreeAngleToDouble(d[0], d[1], d[2]);
        }
        public static double ConvertDegreeAngleToDouble(double degrees, double minutes, double seconds)
        {
            //Decimal degrees = 
            //   whole number of degrees, 
            //   plus minutes divided by 60, 
            //   plus seconds divided by 3600

            return degrees + (minutes / 60) + (seconds / 3600);
        }
    }
}
