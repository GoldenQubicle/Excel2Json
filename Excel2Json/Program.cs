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

            
            RegexFilter parse = new RegexFilter();
            parse.LoadJson();
            parse.Filter();

            //// stuffies below needed for raw import and/or later parsing by regex filter
            //// currently all disabled since working with single json

            //String Filename = "C:\\Users\\Erik\\Desktop\\MAAND 1, WEEK 2\\MAAND 1, WEEK 2, DAG 1\\WOORDZOEKER\\WZ makkelijk LANDBOUW.xlsx";
            //Dictionary<string, List<string>> content = new Dictionary<string, List<string>>();
            //List<string> contentRaw;
            //Import import = new Import();
            //Export export = new Export();


            //// read entire sheet 
            //contentRaw = import.Read(Filename);
            //export.SaveRaw(contentRaw);


            //// more nimble parsin, however, prolly obsolete once regex filtering is in place
            //content.Add("letters", import.LetterArray(Filename));
            //content.Add("words", import.WordArray(Filename));         
            //export.Save(content);

            Console.Read();
        }
    }
}


