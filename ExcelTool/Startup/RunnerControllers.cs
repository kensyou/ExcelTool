using AddinX.Bootstrap.Contract;
using Autofac;
using ExcelTool.Helper;
using ExcelTool.Modules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.Startup
{
    internal class RunnerControllers : IRunner
    {
        public void Execute(IRunnerMain bootstrap)
        {
            var thisAssembly = Assembly.GetExecutingAssembly();

            var bootstrapper = (Bootstrapper)bootstrap;
            bootstrapper?.Builder.RegisterAssemblyTypes(thisAssembly)
                .Where(t => t.Name.ToLower().EndsWith("controller"));


            bootstrapper?.Builder.RegisterType<ExcelHelper>().As<IExcelHelper>().SingleInstance();
            bootstrapper?.Builder.RegisterType<ZenrinModule>().As<IZenrinModule>().SingleInstance();
        }
    }
}
