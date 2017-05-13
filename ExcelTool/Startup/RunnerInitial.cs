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
using Prism.Events;

namespace ExcelTool.Startup
{
    internal class RunnerInitial : IRunner
    {
        public void Execute(IRunnerMain bootstrap)
        {
            var bootstrapper = (Bootstrapper)bootstrap;

            // Excel Application
            bootstrapper?.Builder.RegisterInstance(AddinContext.ExcelApp).ExternallyOwned();

            // Ribbon
            bootstrapper?.Builder.RegisterInstance(new AddinRibbon());

            // ILogger
            bootstrapper?.Builder.RegisterInstance<ILogger>(new SerilogLogger());

            // Event Aggregator
            bootstrapper?.Builder.RegisterInstance<IEventAggregator>(new EventAggregator());

        }
    }
}
