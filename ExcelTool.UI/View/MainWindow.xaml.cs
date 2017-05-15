using BruTile.Predefined;
using ExcelTool.UI.ViewModel;
using Mapsui;
using Mapsui.Layers;
using Mapsui.Projection;
using Mapsui.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelTool.UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        readonly MainWindowViewModel _Model;
        public MainWindow(MainWindowViewModel model)
        {
            InitializeComponent();
            DataContext = model;
            _Model = model;
            MyMapControl.RenderMode = Mapsui.UI.Wpf.RenderMode.Wpf;
            MyMapControl.Map = model.Map;
            HighwayList.SelectionChanged += ReloadMap;
            StyleList.SelectionChanged += ReloadMap;
        }

        private void ReloadMap(object sender, SelectionChangedEventArgs e)
        {
            MyMapControl.Map.Layers.Clear();
            MyMapControl.Map = _Model.Map;
            MyMapControl.Refresh();
        }
    }
}
