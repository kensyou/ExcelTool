using AddinX.Logging;
using AddinX.Wpf.Contract;
using Autofac;
using ExcelTool.Modules;
using ExcelTool.UI;
using ExcelTool.UI.ViewModel;
using ExcelTool.ZenrinIC.Models;
using Mapsui;
using Mapsui.Geometries;
using Mapsui.Layers;
using Mapsui.Projection;
using Mapsui.Providers;
using Mapsui.Styles;
using Mapsui.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace ExcelTool.Controller
{
    public class SampleController : IDisposable
    {
        private readonly ILogger logger;
        private IWpfHelper wpfHelper;
        private IZenrinModule _Zenrin;
        public SampleController(ILogger logger, IWpfHelper wpfHelper, IZenrinModule zenrin)
        {
            this.logger = logger;
            this.wpfHelper = wpfHelper;
            this._Zenrin = zenrin;
        }

        public void OpenForm()
        {
            logger.Debug("Inside show message method");
            //var window = AddinContext.Container.Resolve<MainWindow>();
            //wpfHelper.Show(window);
            var thread = new Thread(() =>
            {

                try
                {
                    var wpfViewModel = AddinContext.Container.Resolve<MainWindowViewModel>();
                    wpfViewModel.MapGenerator = _Zenrin.MapGenerator;
                    wpfViewModel.Highways = new ObservableCollection<Highway>(_Zenrin.Highways);
                    wpfViewModel.SelectedHighway = wpfViewModel.Highways[0];
                    //wpfViewModel.Map = _Zenrin.CreateMap();

                    var wpfWindow = AddinContext.Container.Resolve<MainWindow>();
                    wpfWindow.Show();
                    wpfWindow.Closed += (sender2, e2) => wpfWindow.Dispatcher.InvokeShutdown();

                    Dispatcher.Run();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        public void Dispose()
        {
            wpfHelper = null;
        }

    }

}
