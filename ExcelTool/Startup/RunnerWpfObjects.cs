using AddinX.Bootstrap.Contract;
using AddinX.Logging;
using ILogger = AddinX.Logging.ILogger;
using AddinX.Wpf.Contract;
using AddinX.Wpf.Implementation;
using Autofac;
using ExcelTool.Manipulation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTool.UI;
using ExcelTool.UI.ViewModel;
using ExcelTool.Helper;
using ExcelTool.Modules;

namespace ExcelTool.Startup
{
    internal class RunnerWpfObjects : IRunner
    {
        public void Execute(IRunnerMain bootstrap)
        {
            var bootstrapper = (Bootstrapper)bootstrap;

            bootstrapper?.Builder.RegisterType<MainWindow>();
            bootstrapper?.Builder.RegisterType<MainWindowViewModel>();
            
        }
    }
}
