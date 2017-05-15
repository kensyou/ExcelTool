using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using NetOffice.ExcelApi;
using Mapsui;
using ExcelTool.ZenrinIC.Models;
using ExcelTool.UI.Model;

namespace ExcelTool.Modules
{

    public interface IZenrinModule
    {
        Func<Highway, MapStyle, Map> MapGenerator { get; }

        IEnumerable<Highway> Highways { get; }
        Task ImportInterchangeData();
        Task ExportInterchangeDataAsMergeSql();
    }
}
