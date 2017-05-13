using AddinX.Configuration.Contract;
using Autofac;
using ExcelTool.Controller;
using NetOffice.ExcelApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelTool
{
    internal static class AddinContext
    {
        private static MainController ctrls;

        private static AddInSettings settings;

        public static AddInSettings Settings => settings ?? (settings = new AddInSettings());

        public static CancellationTokenSource TokenCancellationSource { get; set; }

        public static IContainer Container { get; set; }

        public static Application ExcelApp { get; set; }

        public static IConfigurationManager ConfigManager { get; set; }

        public static MainController MainController => ctrls ?? (ctrls = Container.Resolve<MainController>());
    }
}
