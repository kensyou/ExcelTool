using Microsoft.SqlServer.Types;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.ZenrinIC.Models
{
    public class Interchange
    {
        public int TempInterchangeId { get; set; }
        public string FileName { get; set; }
        public int PrefectureCode { get; set; }
        public int RowNumber { get; set; }
        public string ZenrinTypeCode { get; set; }
        public int PrefectureICSerial { get; set; }
        public string IC_Kana { get; set; }
        public string IC_Kanji { get; set; }
        public string HighwayDisplay { get { return String.Join("/", Highways.Select(s => s.Item1.ToString())); } }
        public List<Tuple<Highway, int>> Highways { get; set; }
        public double Latitude { get; set; }
        public double Longitude { get; set; }
        public DateTime DataDate { get; set; }

        public static IEnumerable<Interchange> NormalizedInterchange(IEnumerable<Highway> roads)
        {
            var result = roads.Aggregate(new List<Interchange>(), (acc, item) =>
            {
                var sortedInterchanges = SortInterchanges(item.Interchanges);

                foreach (var i in sortedInterchanges)
                {
                    var existingIC = acc.FirstOrDefault(s => s.PrefectureCode == i.PrefectureCode && s.IC_Kanji == i.IC_Kanji);

                    if (existingIC == null)
                    {
                        existingIC = new Interchange
                        {
                            FileName = i.FileName,
                            RowNumber = i.RowNumber,
                            PrefectureCode = i.PrefectureCode,
                            ZenrinTypeCode = i.ZenrinTypeCode,
                            PrefectureICSerial = i.PrefectureICSerial,
                            IC_Kana = i.IC_Kana,
                            IC_Kanji = i.IC_Kanji,
                            Latitude = i.Latitude,
                            Longitude = i.Longitude,
                            DataDate = i.DataDate,
                            Highways = new List<Tuple<Highway, int>>()
                        };
                        acc.Add(existingIC);
                    }
                    else
                    {
                        if (existingIC.Highways.First().Item1.HighwayKanji == item.HighwayKanji) continue; // one road cannot have duplicate name IC
                    }
                    existingIC.Highways.Add(Tuple.Create(item, i.SortOrder));
                }
                return acc;
            });
            var index = 1;
            return result.Aggregate(new List<Interchange>(), (acc, item) =>
            {
                item.TempInterchangeId = index++;
                acc.Add(item);
                return acc;
            }).ToList();
        }
        private static SqlGeography tokyoKouKyo = SqlGeography.Point(35.685175, 139.7506108, 4326);
        public static List<InterchangeParsed> SortInterchanges(List<InterchangeParsed> ics)
        {
            var srid = 4326;
            if (ics == null || !ics.Any()) return ics;
            if (ics.Count() == 1)
            {
                ics[0].SortOrder = 1;
                return ics;
            }
            List<Tuple<InterchangeParsed, SqlGeography, double>> allLines = new List<Tuple<InterchangeParsed, SqlGeography, double>>();
            foreach (var startingIC in ics)
            {
                List<SqlGeography> remainingICs = ics.OrderBy(s => s.IC_Kanji == startingIC.IC_Kanji ? 0 : 1).Select(s => SqlGeography.Point(s.Latitude, s.Longitude, srid)).ToList();
                SqlGeography currentIC = remainingICs[0];

                SqlGeographyBuilder Builder = new SqlGeographyBuilder();
                Builder.SetSrid(4326);
                Builder.BeginGeography(OpenGisGeographyType.LineString);
                Builder.BeginFigure((double)currentIC.Lat, (double)currentIC.Long);
                remainingICs.Remove(currentIC);
                // While there are still unvisited cities
                while (remainingICs.Count > 0)
                {
                    remainingICs.Sort(delegate (SqlGeography p1, SqlGeography p2)
                    { return p1.STDistance(currentIC).CompareTo(p2.STDistance(currentIC)); });

                    // Move to the closest destination
                    currentIC = remainingICs[0];

                    // Add this city to the tour route
                    Builder.AddLine((double)currentIC.Lat, (double)currentIC.Long);

                    // Update the list of remaining cities
                    remainingICs.Remove(currentIC);
                }

                // End the geometry
                Builder.EndFigure();
                Builder.EndGeography();

                // Return the constructed geometry
                var resultingLine = Builder.ConstructedGeography;
                allLines.Add(Tuple.Create(startingIC, resultingLine, resultingLine.STLength().Value));
            }
            var rankedRoute = allLines.OrderBy(s => s.Item3).Take(2).ToList();
            var bestRoute = rankedRoute.OrderBy(s => SqlGeography.Point(s.Item1.Latitude, s.Item1.Longitude, 4326).STDistance(tokyoKouKyo)).First().Item2;
            //var bestRoute = (rankedRoute.First()).Item2;

            var sortingBoard = new List<Tuple<string, int>>();
            var numOfPoints = bestRoute.STNumPoints();
            for (var i = 1; i <= numOfPoints; i++)
            {
                var point = bestRoute.STPointN(i);
                sortingBoard.Add(Tuple.Create(point.ToString(), i));
            }

            var result = ics.Aggregate(new List<InterchangeParsed>(), (acc, item) =>
            {
                var f = sortingBoard.First(x => x.Item1 == SqlGeography.Point(item.Latitude, item.Longitude, srid).ToString());
                item.SortOrder = f.Item2;
                acc.Add(item);
                return acc;
            }).ToList();
            return result;
        }


        //public Image GetBoundaryImage(GetBoundaryRequest request)
        //{
        //    MapView mapView = GetMapView(request.Force, request.Code, request.ImageSize);
        //    Image i = new Bitmap(request.ImageSize.Width, request.ImageSize.Height);
        //    Graphics graphics = Graphics.FromImage(i);

        //    SqlGeography boundary = databaseService.GetBoundary(request.Force, request.Code);

        //    Pen boundaryPen = new Pen(Color.Red, 3.0F);
        //    DrawBoundary(mapView, graphics, boundary, boundaryPen);

        //    return i;

        //}


        //private void DrawBoundary(MapView mapView, Graphics graphics, SqlGeography geography, Pen pen)
        //{
        //    if (geography.STNumGeometries() > 1)
        //    {
        //        for (int geom = 1; geom <= geography.STNumGeometries(); geom++)
        //        {
        //            DrawBoundary(mapView, graphics, geography.STGeometryN(geom), pen);
        //        }
        //    }
        //    PixelPointArraySink pixelPointArraySink = new PixelPointArraySink(mapView);
        //    geography.Populate(pixelPointArraySink);
        //    foreach (Point[] shape in pixelPointArraySink.Shapes)
        //    {
        //        graphics.DrawPolygon(pen, shape);
        //    }
        //}

    }
}
