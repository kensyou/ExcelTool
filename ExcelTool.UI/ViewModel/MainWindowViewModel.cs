using ExcelTool.UI.Model;
using ExcelTool.ZenrinIC.Models;
using Mapsui;
using Mapsui.Geometries;
using Mapsui.Layers;
using Mapsui.Projection;
using Mapsui.Providers;
using Mapsui.Styles;
using Mapsui.Utilities;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.UI.ViewModel
{
    public class MainWindowViewModel : BindableBase, INotifyPropertyChanged
    {

        public MainWindowViewModel()
        {
            MapStyles = new ObservableCollection<MapStyle>(MapStyle.GetMapboxStyles());
        }
        public event PropertyChangedEventHandler PropertyChanged;
        // This method is called by the Set accessor of each property.
        // The CallerMemberName attribute that is applied to the optional propertyName
        // parameter causes the property name of the caller to be substituted as an argument.
        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public ObservableCollection<MapStyle> MapStyles { get; set; }
        private MapStyle _SelectedMapStyle;

        public MapStyle SelectedMapStyle
        {
            get { return _SelectedMapStyle; }
            set
            {
                _SelectedMapStyle = value;
                NotifyPropertyChanged();
                Map = MapGenerator(null, value);
            }
        }

        public ObservableCollection<Highway> Highways { get; set; }

        private Highway _SelectedHighway;

        public Highway SelectedHighway
        {
            get { return _SelectedHighway; }
            set
            {
                _SelectedHighway = value;
                NotifyPropertyChanged();
                Map = MapGenerator(value, null);
            }
        }
        public Func<Highway, MapStyle, Map> MapGenerator { get; set; }
        public Map Map { get; set; }

    }

}
