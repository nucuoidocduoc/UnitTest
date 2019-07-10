using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace WpfApp1
{
    public class DependencyPropertyInfo
    {
        private DependencyPropertyDescriptor _descriptor;
        private DependencyObject _element;

        public DependencyPropertyInfo(DependencyPropertyDescriptor descriptor, DependencyObject element)
        {
            this._descriptor = descriptor;
            this._element = element;
            PropertyInfo propertyInfo = descriptor.ComponentType.GetProperty(descriptor.DependencyProperty.Name);
            var att = propertyInfo.GetCustomAttributes(true);
        }

        public Type PropertyType
        {
            get { return _descriptor.PropertyType; }
        }

        public string Name
        {
            get { return _descriptor.Name; }
        }

        public object Value
        {
            get { return _element.GetValue(_descriptor.DependencyProperty); }
            set { _element.SetValue(_descriptor.DependencyProperty, value); }
        }
    }
}