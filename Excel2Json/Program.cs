using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace Excel2Json
{
    class Program
    {            
        static void Main(string[] args)
        {
            Import import = new Import();
            import.Read("C:\\Users\\Erik\\Desktop\\TESTMAAND 1, WEEK 2 14 MEI VOLLEDIG\\MAAND 1, WEEK 2\\MAAND 1, WEEK 2, DAG 2\\WOORDZOEKER\\WZ makkelijk FRUIT.xlsx");

            Console.Read();
        }
    }
}


