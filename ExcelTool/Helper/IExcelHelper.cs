using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace ExcelTool.Helper
{
    public interface IExcelHelper
    {
        void ValidateWorkbookNotInProtectedView();
        void ChangeMousePointer(XlMousePointer pt);
        void RestoreDefaultState();
        void SwitchToBusyState();
        Worksheet AddWorksheetIfExistsShowMessage(string newWorksheetName, string dispName);
        Worksheet AddWorksheetDuplicate(string newWorksheetName);
        Worksheet AddWorksheetIfNotExists(string newWorksheetName);
        Worksheet AddAfterWorksheetIfNotExists(string newWorksheetName);
        Task DownloadDataOnExcel<T>(string sheetName, Func<Task<IEnumerable<T>>> dataGetter, Func<Worksheet, IEnumerable<T>, Worksheet> excelWriter);
        Task DownloadDataOnExcel<T, U>(string sheetName, U para1, Func<U, Task<IEnumerable<T>>> dataGetter, Func<Worksheet, IEnumerable<T>, Worksheet> excelWriter);
        Task DownloadDataOnExcel<T, U, V>(string sheetName, U para1, V para2, Func<U, V, Task<IEnumerable<T>>> dataGetter, Func<Worksheet, IEnumerable<T>, Worksheet> excelWriter);
    }
}
