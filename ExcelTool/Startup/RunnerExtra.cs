using AddinX.Bootstrap.Contract;
using AddinX.Wpf.Contract;
using AddinX.Wpf.Implementation;
using Autofac;
using ExcelTool.Manipulation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.Startup
{
    internal class RunnerExtra : IRunner
    {
        public void Execute(IRunnerMain bootstrap)
        {
            var bootstrapper = (Bootstrapper)bootstrap;
            bootstrapper?.Builder.RegisterType<ExcelInteraction>();

            bootstrapper?.Builder.RegisterType<ExcelDnaWpfHelper>().As<IWpfHelper>();
        }
    }
}
