using ExcelDna.Integration;
using NetOffice.ExcelApi.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xl = NetOffice.ExcelApi;
namespace ExcelTool.Helper
{
    public class ExcelHelper : IExcelHelper
    {
        private Xl.Application _App;
        public ExcelHelper()
        {
            _App = AddinContext.ExcelApp;
        }
        
        public void ValidateWorkbookNotInProtectedView()
        {
            if (_App.ActiveProtectedViewWindow != null)
            {
                throw new Exception("Please enable edit mode for the active workbook or create a new workbook before proceeding.");
            }
        }
        public void RestoreDefaultState()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                _App.StatusBar = String.Empty;
                _App.ScreenUpdating = true;
                _App.Cursor = XlMousePointer.xlDefault;
            });
        }
        public void ChangeMousePointer(XlMousePointer pt)
        {
            // Cursor need to be queued as macro
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                _App.Cursor = pt;
            });
        }

        public void SwitchToBusyState()
        {
            ChangeMousePointer(XlMousePointer.xlWait);
            _App.ScreenUpdating = false;
        }
        private DialogResult ShowMessagebox(string msg, string dispName)
        {
            return MessageBox.Show(String.Join("\r\n", msg), dispName, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
        }
        public Xl.Worksheet AddWorksheetIfExistsShowMessage(string newWorksheetName, string dispName)
        {
            var wb = _App.ActiveWorkbook ?? _App.Workbooks.Add();
            var wss = wb.Worksheets;
            var wsNames = new List<String>();
            foreach (dynamic s in wss)
            {
                wsNames.Add(s.Name);
            }
            if (wsNames.Contains(newWorksheetName))
            {
                var msg = $"Already have worksheet named {newWorksheetName}. \r\n Do you want to download this master?";
                if (ShowMessagebox(msg, dispName) == DialogResult.No) throw new Exception();
                return AddWorksheetDuplicate(newWorksheetName);
            }

            var ws = (Xl.Worksheet)wss.Add();
            ws.Name = newWorksheetName;
            return ws;
        }
        public Xl.Worksheet AddWorksheetDuplicate(string newWorksheetName)
        {
            var wb = _App.ActiveWorkbook ?? _App.Workbooks.Add();
            var wss = wb.Worksheets;
            var wsNames = new List<String>();
            var count = 1;
            foreach (dynamic s in wss)
            {
                if (s.Name.IndexOf(newWorksheetName) != -1) count++;
            }
            var ws = (Xl.Worksheet)wss.Add();
            ws.Name = $"{newWorksheetName} {(count == 1 ? "" : $"({count})")}";
            return ws;
        }
        public Xl.Worksheet AddWorksheetIfNotExists(string newWorksheetName)
        {
            var wb = _App.ActiveWorkbook ?? _App.Workbooks.Add();
            var wss = wb.Worksheets;
            var wsNames = new List<String>();
            foreach (dynamic s in wss)
            {
                wsNames.Add(s.Name);
            }
            if (wsNames.Contains(newWorksheetName)) throw new Exception($"Already have worksheet named {newWorksheetName}");

            Xl.Worksheet ws = (Xl.Worksheet)wss.Add();
            ws.Name = newWorksheetName;
            return ws;
        }
        public Xl.Worksheet AddAfterWorksheetIfNotExists(string newWorksheetName)
        {
            var wb = _App.ActiveWorkbook ?? _App.Workbooks.Add();
            var wss = wb.Worksheets;
            var wsNames = new List<String>();
            foreach (dynamic s in wss)
            {
                wsNames.Add(s.Name);
            }
            if (wsNames.Contains(newWorksheetName)) throw new Exception($"Already have worksheet named {newWorksheetName}");

            var ws = (Xl.Worksheet)wss.Add(null, after: _App.ActiveSheet);
            ws.Name = newWorksheetName;
            return ws;
        }

        public async Task DownloadDataOnExcel<T>(string sheetName, Func<Task<IEnumerable<T>>> dataGetter, Func<Xl.Worksheet, IEnumerable<T>, Xl.Worksheet> excelWriter)
        {
            try
            {
                ChangeMousePointer(XlMousePointer.xlWait);
                _App.ScreenUpdating = false;

                var dispName = $"{sheetName} Master";
                var noDataMsg = $"No {dispName}";

                var data = await dataGetter();
                if (!data.Any()) throw new Exception(noDataMsg);

                var ws = AddWorksheetIfExistsShowMessage(sheetName, dispName);
                excelWriter(ws, data).Select();
            }
            finally
            {
                RestoreDefaultState();
            }
        }
        public async Task DownloadDataOnExcel<T, U>(string sheetName, U para1, Func<U, Task<IEnumerable<T>>> dataGetter, Func<Xl.Worksheet, IEnumerable<T>, Xl.Worksheet> excelWriter)
        {
            try
            {
                ChangeMousePointer(XlMousePointer.xlWait);
                _App.ScreenUpdating = false;

                var dispName = $"{sheetName} {para1}";
                var sheetNameWithYM = $"{sheetName}_{para1}";
                var noDataMsg = $"No {dispName}";

                var data = await dataGetter(para1);
                if (!data.Any()) throw new Exception(noDataMsg);

                var ws = AddWorksheetIfExistsShowMessage(sheetNameWithYM, dispName);
                excelWriter(ws, data).Select();
            }
            finally
            {
                RestoreDefaultState();
            }
        }

        public async Task DownloadDataOnExcel<T, U, V>(string sheetName, U para1, V para2, Func<U, V, Task<IEnumerable<T>>> dataGetter, Func<Xl.Worksheet, IEnumerable<T>, Xl.Worksheet> excelWriter)
        {
            try
            {
                ChangeMousePointer(XlMousePointer.xlWait);
                _App.ScreenUpdating = false;

                var dispName = $"{sheetName} {para1}/{String.Format("{0:D2}", para2)}";
                var sheetNameWithYM = $"{sheetName}_{para1}_{para2}";
                var noDataMsg = $"No {dispName}";

                var data = await dataGetter(para1, para2);
                if (!data.Any()) throw new Exception(noDataMsg);

                var ws = AddWorksheetIfExistsShowMessage(sheetNameWithYM, dispName);
                excelWriter(ws, data).Select();
            }
            finally
            {
                RestoreDefaultState();
            }
        }

    }
}
