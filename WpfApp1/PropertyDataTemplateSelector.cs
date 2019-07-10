using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace WpfApp1
{
    public class PropertyDataTemplateSelector : DataTemplateSelector
    {
        public DataTemplate DefaultnDataTemplate { get; set; }
        public DataTemplate BooleanDataTemplate { get; set; }
        public DataTemplate EnumDataTemplate { get; set; }

        public override DataTemplate SelectTemplate(object item,
                   DependencyObject container)
        {
            MyData data = item as MyData;
            if (data.IsRuleSet) {
                return DefaultnDataTemplate;
            }
            else {
                return BooleanDataTemplate;
            }
        }
    }
}