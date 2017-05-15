using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.UI.Model
{ 
    public class MapStyle
    {
        public string Name { get; set; }

        public static MapStyle[] GetMapboxStyles()
        {
            var styles = new string[]
            {
                "mapbox.streets",
                "mapbox.light",
                "mapbox.dark",
                "mapbox.satellite",
                "mapbox.streets-satellite",
                "mapbox.wheatpaste",
                "mapbox.streets-basic",
                "mapbox.comic",
                "mapbox.outdoors",
                "mapbox.run-bike-hike",
                "mapbox.pencil",
                "mapbox.pirates",
                "mapbox.emerald",
                "mapbox.high-contrast",
            };
            return styles.Select(s => new MapStyle { Name = s }).ToArray();
        }
    }
}
