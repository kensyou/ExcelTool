using AddinX.Logging;
using AddinX.Wpf.Contract;
using Autofac;
using ExcelTool.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.Controller
{
    public class SampleController : IDisposable
    {
        private readonly ILogger logger;
        private IWpfHelper wpfHelper;

        public SampleController(ILogger logger, IWpfHelper wpfHelper)
        {
            this.logger = logger;
            this.wpfHelper = wpfHelper;
        }

        public void OpenForm()
        {
            logger.Debug("Inside show message method");
            var window = AddinContext.Container.Resolve<MainWindow>();
            wpfHelper.Show(window);
        }

        public void Dispose()
        {
            wpfHelper = null;
        }
    }
}
