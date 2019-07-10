using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1
{
    public class ViewModel : BindableBase
    {
        private ObservableCollection<MyData> _dataSet;
        public ObservableCollection<MyData> DataSet { get => _dataSet; set => SetProperty(ref _dataSet, value); }

        public ViewModel()
        {
            _dataSet = new ObservableCollection<MyData>();
            _dataSet.Add(new MyData("nguyen", false));
            _dataSet.Add(new MyData("phuong", true));
            _dataSet.Add(new MyData("nguyen", false));
            _dataSet.Add(new MyData("phuong", true));
            _dataSet.Add(new MyData("nguyen", false));
            _dataSet.Add(new MyData("phuong", true));
            _dataSet.Add(new MyData("nguyen", false));
            _dataSet.Add(new MyData("phuong", true));
            _dataSet.Add(new MyData("nguyen", false));
            _dataSet.Add(new MyData("phuong", true));
            _dataSet.Add(new MyData("nguyen", false));
            _dataSet.Add(new MyData("phuong", true));
            _dataSet.Add(new MyData("nguyen", false));
            _dataSet.Add(new MyData("phuong", true));
            _dataSet.Add(new MyData("nguyen", false));
            _dataSet.Add(new MyData("phuong", true));
            _dataSet.Add(new MyData("nguyen", false));
            _dataSet.Add(new MyData("phuong", true));
        }
    }
}