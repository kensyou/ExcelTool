using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.Modules
{
    public interface IIndustrialModule
    {
        Task ParseIndustrialStackingList();
    }
}
