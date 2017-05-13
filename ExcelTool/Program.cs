using ExcelDna.Integration;
using ExcelDna.Logging;
using ExcelTool.Startup;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xl = NetOffice.ExcelApi;
namespace ExcelTool
{
    public class Program: IExcelAddIn
    {
        public void AutoClose()
        {
            throw new NotImplementedException();
        }

        public void AutoOpen()
        {
            try
            {
                // Token cancellation is useful to close all existing Tasks<> before leaving the application
                AddinContext.TokenCancellationSource = new CancellationTokenSource();

                AddinContext.ExcelApp = new Xl.Application(null, ExcelDnaUtil.Application);
                
                // Start the bootstrapper now
                new Bootstrapper(AddinContext.TokenCancellationSource.Token).Start();
            }
            catch(Exception e)
            {
                LogDisplay.RecordLine(e.Message);
                LogDisplay.RecordLine(e.StackTrace);
                LogDisplay.Show();
            }
        }
    }
}
