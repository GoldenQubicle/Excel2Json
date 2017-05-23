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
            JSONHelper jsonHelper = new JSONHelper();
            Parser parser = new Parser();
            CalcWordPos calcWordPos = new CalcWordPos();
            List<string> content = new List<string>();
            Dictionary<string, List<string>> contentFormatted = new Dictionary<string, List<string>>();
            String filename = "C:\\Users\\Erik\\Desktop\\MAAND 1, WEEK 2\\MAAND 1, WEEK 2, DAG 1\\WOORDZOEKER\\WZ middel LANDBOUW.xlsx";

            // proper routine needs to go here =) 
            //content = import.Read(Filame);
            
            // temp routine to save single excel sheet as json 
            //export.SaveRaw(import.Read(filename));
            
            // read said json as list
            content = parser.scrubContent(jsonHelper.loadRaw());

            jsonHelper.getRowsColumns().ToList().ForEach(x => contentFormatted.Add(x.Key, x.Value));
            parser.getWordsLetters(content).ToList().ForEach(x => contentFormatted.Add(x.Key, x.Value));
            parser.getSolutionInfo(content).ToList().ForEach(x => contentFormatted.Add(x.Key, x.Value));

            calcWordPos.getCol(contentFormatted);

            //export.SaveIntermediate(contentFormatted);
            Console.Read();
        }
    }
}


