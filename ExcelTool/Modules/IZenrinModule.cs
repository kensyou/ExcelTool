using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using NetOffice.ExcelApi;

namespace ExcelTool.Modules
{

    public interface IZenrinModule
    {
        Task ImportInterchangeData();
    }
}
