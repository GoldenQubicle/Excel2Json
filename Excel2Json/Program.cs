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
            //Dictionary<string, List<string>> contentRaw = new Dictionary<string, List<string>>();

            Dictionary<string, string> levels = Levels();

            // proper routine needs to go here =) 

            string path = @"C:\Users\Erik\Desktop\MAAND 1, WEEK 2";
            //foreach (string file in scraper.getFiles(path))
            //{
            for (int i = 0; i < 4; i++)
            {
                Dictionary<string, List<string>> contentRaw = new Dictionary<string, List<string>>();
                string file = scraper.getFiles(path)[i];
                Console.WriteLine("opening file " + file);
                contentRaw = import.readFile(file, levels);


                //String filename = "C:\\Users\\Erik\\Desktop\\MAAND 1, WEEK 2\\MAAND 1, WEEK 2, DAG 1\\WOORDZOEKER\\WZ moeilijk LANDBOUW.xlsx";

                //contentRaw = import.readFile(filename, levels);
                //string key = "20";

                // temp routine to save single excel sheet raw data as json 
                //export.SaveRaw(contentRaw[key], key);

                // temp: read said json as list and pars it, also temp key
                //List<string> content = new List<string>();
                //content = jsonHelper.loadRaw(key);
                //content = parser.scrubContent(content);

                foreach (string key in contentRaw.Keys)
                {

                    List<string> content = new List<string>();
                    Dictionary<string, List<string>> contentFormattedStrings = new Dictionary<string, List<string>>();
                    Dictionary<string, List<int>> contentFormattedInts = new Dictionary<string, List<int>>();
                    Dictionary<string, int> contentColRow = new Dictionary<string, int>();

                    content = parser.scrubContent(contentRaw[key]);

                    parser.getWordsLetters(content).ToList().ForEach(x => contentFormattedStrings.Add(x.Key, x.Value));
                    parser.getSolutionInfo(content).ToList().ForEach(x => contentFormattedStrings.Add(x.Key, x.Value));

                    calcWordPos.getColRow(key).ToList().ForEach(x => contentColRow.Add(x.Key, x.Value));
                    calcWordPos.FindWords(contentFormattedStrings).ToList().ForEach(x => contentFormattedInts.Add(x.Key, x.Value));

                    calcWordPos.FindSolution(contentFormattedStrings).ToList().ForEach(x => contentFormattedInts.Add(x.Key, x.Value));

                    export.SaveFinal(contentFormattedStrings, contentFormattedInts, contentColRow, key);


                }
            }

            //jsonHelper.getRowsColumns().ToList().ForEach(x => contentColRow.Add(x.Key, x.Value));

            //parser.getWordsLetters(content).ToList().ForEach(x => contentFormattedStrings.Add(x.Key, x.Value));
            //parser.getSolutionInfo(content).ToList().ForEach(x => contentFormattedStrings.Add(x.Key, x.Value));

            //calcWordPos.getColRow(contentColRow);
            //calcWordPos.FindFirstLetter(contentFormattedStrings).ToList().ForEach(x => contentFormattedInts.Add(x.Key, x.Value));

            //export.SaveFinal(contentFormattedStrings, contentFormattedInts, contentColRow);

            //export.SaveIntermediate(contentFormattedStrings);



            Console.Read();
        }

    public static Dictionary<string, string> Levels()
    {
        Dictionary<string, string> levelSublevel = new Dictionary<string, string>();

        string[] levels = { "makkelijk ", "middel ", "moeilijk ", "moeilijk + " };

        string[] sublevels = { "(makkelijk)", "(middel)", "(moeilijk)" };

        foreach (string level in levels)
        {
            foreach (string sublevel in sublevels)
            {
                levelSublevel.Add(level + sublevel, Array.FindIndex(levels, row => row.Contains(level)).ToString() + Array.FindIndex(sublevels, row => row.Contains(sublevel)).ToString());
            }

        }
        return levelSublevel;
    }
}
}


