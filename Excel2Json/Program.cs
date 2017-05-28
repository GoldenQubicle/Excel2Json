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
            factor out retrieve index of first letter words since thats needed to:
            work out the logic for search vertical and diagonal
            DONT FORGET: also need to get the positions for words in solution 
            and speaking off: 
            adjust the js in project to accomodate solution split into words (.e.g join together on client and pass into existing linesolution var)
            
            when above is done then:
            determine themeDay on folder / path
            determine level on filename
            determine sublevel on sheetname               
            
            */

            Import import = new Import();
            Scraper scraper = new Scraper();
            Export export = new Export();
            JSONHelper jsonHelper = new JSONHelper();
            Parser parser = new Parser();
            CalcWordPos calcWordPos = new CalcWordPos();
            Dictionary<string, List<string>> contentRaw = new Dictionary<string, List<string>>();
            

            // proper routine needs to go here =) 

            //string path = @"C:\Users\Erik\Desktop\MAAND 1, WEEK 2";
            //foreach (string file in scraper.getFiles(path))
            //{
            //    Console.WriteLine(file);
            //};

            String filename = "C:\\Users\\Erik\\Desktop\\MAAND 1, WEEK 2\\MAAND 1, WEEK 2, DAG 1\\WOORDZOEKER\\WZ makkelijk LANDBOUW.xlsx";
            contentRaw = import.readFile(filename);

            foreach(string key in contentRaw.Keys)
            {
                List<string> content = new List<string>();
                Dictionary<string, List<string>> contentFormattedStrings = new Dictionary<string, List<string>>();
                Dictionary<string, List<int>> contentFormattedInts = new Dictionary<string, List<int>>();
                Dictionary<string, int> contentColRow = new Dictionary<string, int>();

                content = parser.scrubContent(contentRaw[key]);

                parser.getWordsLetters(content).ToList().ForEach(x => contentFormattedStrings.Add(x.Key, x.Value));
                parser.getSolutionInfo(content).ToList().ForEach(x => contentFormattedStrings.Add(x.Key, x.Value));

                calcWordPos.getColRow(key).ToList().ForEach(x => contentColRow.Add(x.Key, x.Value));

                // so yeah need to pass in the key to determine search directions
                calcWordPos.FindFirstLetter(contentFormattedStrings).ToList().ForEach(x => contentFormattedInts.Add(x.Key, x.Value));

                export.SaveFinal(contentFormattedStrings, contentFormattedInts, contentColRow, key);

            }

            //import.determineLevels();
            // temp routine to save single excel sheet raw data as json 
            //export.SaveRaw(import.Read(filename));

            // read said json as list
            //content = parser.scrubContent(jsonHelper.loadRaw());

            //jsonHelper.getRowsColumns().ToList().ForEach(x => contentColRow.Add(x.Key, x.Value));

            //parser.getWordsLetters(content).ToList().ForEach(x => contentFormattedStrings.Add(x.Key, x.Value));
            //parser.getSolutionInfo(content).ToList().ForEach(x => contentFormattedStrings.Add(x.Key, x.Value));

            //calcWordPos.getColRow(contentColRow);
            //calcWordPos.FindFirstLetter(contentFormattedStrings).ToList().ForEach(x => contentFormattedInts.Add(x.Key, x.Value));

            //export.SaveFinal(contentFormattedStrings, contentFormattedInts, contentColRow);

            //export.SaveIntermediate(contentFormattedStrings);



            Console.Read();
        }
    }
}


