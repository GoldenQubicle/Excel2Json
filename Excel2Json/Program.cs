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
            String Filename = "C:\\Users\\Erik\\Desktop\\MAAND 1, WEEK 2\\MAAND 1, WEEK 2, DAG 1\\WOORDZOEKER\\WZ makkelijk LANDBOUW.xlsx";
            Dictionary<string, List<string>> content = new Dictionary<string, List<string>>();

            Import import = new Import();
            Export export = new Export();

            content.Add("letters", import.LetterArray(Filename));
            content.Add("words", import.WordArray(Filename));

            export.Save(content);

            //Console.Read();
        }
    }
}


