using AddinX.Bootstrap.Autofac;
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelTool.Startup
{
    internal class Bootstrapper : AutofacRunnerMain
    {
        public Bootstrapper(CancellationToken token)
                    : base(token)
        {
        }

        public override void Start()
        {
            ExcelAsyncUtil.QueueAsMacro(() => base.Start());
        }

        public override void ExecuteAll()
        {
            base.ExecuteAll();
            AddinContext.Container = GetContainer();
        }
    }
}
