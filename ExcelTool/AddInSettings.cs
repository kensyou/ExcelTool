using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    internal class AddInSettings
    {
        public string AddinName => AddinContext.ConfigManager.AppSettings["addin:name"];

        public string AddinPath { get; set; }
    }
}
