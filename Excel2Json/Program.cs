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

            /*
            TODO
            calculate word positions
            determine level on filename
            scrape folder for filenames             
            */

            Import import = new Import();
            Export export = new Export();
            JSONBuilder json = new JSONBuilder();
            RegexFilter parse = new RegexFilter();

            String Filename = "C:\\Users\\Erik\\Desktop\\MAAND 1, WEEK 2\\MAAND 1, WEEK 2, DAG 1\\WOORDZOEKER\\WZ makkelijk LANDBOUW.xlsx";

            json.SplitScrubbedList(parse.Filter(import.Read(Filename)));

            //parse.LoadJson(); // temp to bypass using excel assembly

            export.Save(json.Build());

            Console.Read();
        }
    }
}


