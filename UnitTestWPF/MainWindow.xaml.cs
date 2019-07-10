using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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

namespace UnitTestWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public MainWindow()
        {
            _fieldsDisplay = new ObservableCollection<ScheduleFieldDisplay>();
            FieldsDisplay = new ObservableCollection<ScheduleFieldDisplay>() { new ScheduleFieldDisplay() { FieldHeading="nguyen",IsUpdate=true},
            new ScheduleFieldDisplay() { FieldHeading="nguyen",IsUpdate=true},
            new ScheduleFieldDisplay() { FieldHeading="trong",IsUpdate=true},
            new ScheduleFieldDisplay() { FieldHeading="phuong",IsUpdate=true},
            new ScheduleFieldDisplay() { FieldHeading="que",IsUpdate=true},
            new ScheduleFieldDisplay() { FieldHeading="bac",IsUpdate=true},
            new ScheduleFieldDisplay() { FieldHeading="ninh",IsUpdate=true}};
            InitializeComponent();
            this.DataContext = this;
        }

        private ObservableCollection<ScheduleFieldDisplay> _fieldsDisplay = new ObservableCollection<ScheduleFieldDisplay>();

        public ObservableCollection<ScheduleFieldDisplay> FieldsDisplay
        {
            get => _fieldsDisplay;
            set
            {
                _fieldsDisplay = value;
                //OnPropertyChanged(nameof(FieldsDisplay));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        //protected virtual void OnPropertyChanged(string propertyName)
        //{
        //    PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        //}
    }
}