using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1
{
    public class MyData
    {
        public bool IsRuleSet { get; set; }
        public string Name { get; set; }

        public MyData(string name, bool isRuleSet)
        {
            Name = name;
            IsRuleSet = isRuleSet;
        }
    }
}