using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System.IO;

namespace Excel2Json
{
    class Program
    {

        static void Main(string[] args)
        {

            /*
             
            TODO 
                                             
            adjust the javascript in project to accomodate solution split into words (i.e. join together on client and pass into existing linesolution var)
            
            check interpunction solution words, ask Bardo & Paul how this was handled in the javascript
            
            determine themeDay on folder / path               

            for extra bonus fun points: select search direction on level, i.e. currently searches vert & dia when its not necessarry
            
            */

            JSONHelper jsonHelper = new JSONHelper();
            Dictionary<string, string> levels = Levels();

            // ~~~~ debug dev stuff ~~~~
            //String filename = "C:\\Users\\Erik\\Desktop\\MAAND 1, WEEK 2\\MAAND 1, WEEK 2, DAG 1\\WOORDZOEKER\\WZ moeilijk LANDBOUW.xlsx";
            //Dictionary<string, List<string>> contentRaw = new Dictionary<string, List<string>>();
            //contentRaw = Import.readFile(filename, levels); 
            //string key = "22"; // note key needs to correspond to sheet # during import
            //ProcessSheet(key, contentRaw);

            //// temp routine to save single excel sheet raw data as json
            // Export.SaveRaw(contentRaw[key], key);

            //// temp: read said json as list and pars it, also temp key
            // List<string> content = new List<string>();
            // content = jsonHelper.loadRaw(key);
            // content = Parser.scrubContent(content);

            // ~~~~ proper routine ~~~~
            string path = @"C:\Users\Erik\Desktop\KK cloud content";
            string savePath = @"C:\Users\Erik\Documents\GitHub\Content\wordsearch\";
            List<string> files = Scraper.getFiles(path);
            int fileCount = 1;
            int dayCount = 1;
            int weekCount = 1;


            // ok so here's a question, why doesnt moeilijk + middel sheet return a proper index all together?!
            for (int i = 1; i < 4; i++)
            {

           //foreach (string file in files) // process everything! 
           //{
                string file = Scraper.getFiles(path)[i];
                string newFolder = "week_" + weekCount + "_day_" + dayCount;
                string folderPath = Scraper.folders[dayCount].ToString();
                if (fileCount > 4)
                {
                    fileCount = 1;
                    dayCount += 1;
                }
                if (dayCount > 5)
                {
                    weekCount += 1;
                    dayCount = 1;
                }
                fileCount += 1;

                Console.WriteLine("opening file " + file);
                Dictionary<string, List<string>> contentRaw = new Dictionary<string, List<string>>();
                contentRaw = Import.readFile(file, levels);
                
                foreach (string key in contentRaw.Keys)
                {
                    ProcessSheet(key, contentRaw, savePath + newFolder);
                }
            }
            Console.Read();
        }

        public static void ProcessSheet(string key, Dictionary<string, List<string>> contentRaw, string folderPath)
        {
            List<string> content = new List<string>();
            Dictionary<string, List<string>> contentFormattedStrings = new Dictionary<string, List<string>>();
            Dictionary<string, List<int>> contentFormattedInts = new Dictionary<string, List<int>>();
            Dictionary<string, int> contentColRow = new Dictionary<string, int>();

            content = Parser.scrubContent(contentRaw[key]);

            Parser.getWordsLetters(content).ToList().ForEach(x => contentFormattedStrings.Add(x.Key, x.Value));
            Parser.getSolutionInfo(content).ToList().ForEach(x => contentFormattedStrings.Add(x.Key, x.Value));

            CalcWordPos.getColRow(key).ToList().ForEach(x => contentColRow.Add(x.Key, x.Value));
            CalcWordPos.FindWords(contentFormattedStrings).ToList().ForEach(x => contentFormattedInts.Add(x.Key, x.Value));
            CalcWordPos.FindSolution(contentFormattedStrings).ToList().ForEach(x => contentFormattedInts.Add(x.Key, x.Value));

            Export.SaveFinal(contentFormattedStrings, contentFormattedInts, contentColRow, key, folderPath);

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


