using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace WpfApp1
{
    public static class TypeHelper
    {
        public static IEnumerable<DependencyPropertyInfo> GetDependencyProperties<T>(DependencyObject element)
        {
            foreach (PropertyDescriptor pd in TypeDescriptor.GetProperties(typeof(T),
                        new Attribute[] { new PropertyFilterAttribute(PropertyFilterOptions.SetValues |
                                                                                PropertyFilterOptions.UnsetValues |
                                                                                PropertyFilterOptions.Valid) })) {
                DependencyPropertyDescriptor dpd = DependencyPropertyDescriptor.FromProperty(pd);
                if (dpd != null) {
                    yield return new DependencyPropertyInfo(dpd, element);
                }
            }
        }
    }
}