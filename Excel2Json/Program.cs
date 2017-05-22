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
            String Filename = "C:\\Users\\Erik\\Desktop\\MAAND 1, WEEK 2\\MAAND 1, WEEK 2, DAG 2\\WOORDZOEKER\\WZ makkelijk FRUIT.xlsx";
            Dictionary<string, List<string>> content;

            Import import = new Import();
            Export export = new Export();

            content = import.LetterArray(Filename);

            export.Save(content);

            //Console.Read();
        }
    }
}


