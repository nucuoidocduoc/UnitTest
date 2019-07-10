using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTestExportExcel
{
    public class Data
    {
        public List<string> Values { get; set; }

        public Data()
        {
            Values = new List<string>();
            for (int i = 0; i < 20; i++) {
                Values.Add(i.ToString());
            }
        }
    }
}