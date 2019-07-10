using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTestExportExcel
{
    public class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("start init data :" + DateTime.Now.ToString("h:mm:ss tt") + "\n");

            Report report = new Report();
            Console.WriteLine("end init data :" + DateTime.Now.ToString("h:mm:ss tt") + "\n");
            Console.WriteLine("start write excel :" + DateTime.Now.ToString("h:mm:ss tt") + "\n");
            //report.CreateExcel("D:\\TestExcel.xlsx");
            WriteDataToExcelSAX.WriteDataSAX();

            Console.WriteLine("end write excel :" + DateTime.Now.ToString("h:mm:ss tt") + "\n");
            Console.ReadKey();
        }
    }
}