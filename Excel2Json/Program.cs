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
            also, prolly want to properly close COM object when scraping entire folders . . =) 
            ALSO CHECMK FOR THE DARN FRIGGIN IJ!
            */

            //Import import = new Import();
            Export export = new Export();
            JSONBuilder json = new JSONBuilder();
            RegexFilter parse = new RegexFilter();
            List<string> content = new List<string>();
            String Filename = "C:\\Users\\Erik\\Desktop\\MAAND 1, WEEK 2\\MAAND 1, WEEK 2, DAG 1\\WOORDZOEKER\\WZ middel LANDBOUW.xlsx";

            //content = import.Read(Filame);
            //export.Save(json.Build());


            // temp routine
            // save single sheet as json 
            //export.SaveRaw(import.Read(Filename));
            // read said json 
            content = parse.ScrubContent(json.LoadRaw());            
            export.SaveIntermediate(parse.getSolutionInfo(content));
             Console.Read();
        }
    }
}


