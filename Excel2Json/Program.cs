﻿using System;
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
            JSONHelper jsonHelper = new JSONHelper();
            Dictionary<string, string> levels = Levels();

            // ~~~~ proper routine ~~~~
            string path = @"C:\Users\Erik\Desktop\KK cloud content";
            //string savePath = @"C:\Users\Erik\Documents\GitHub\Content\wordsearch\";
            string savePath = @"C:\Users\Erik\Documents\GitHub\Content\test\";
            List<string> files = Scraper.getFiles(path);
            int fileCount = 1;
            int dayCount = 1;
            int weekCount = 1;

            for (int i = 0; i < 1; i++)
            {
                //foreach (string file in files) // process everything! 
                                               //{
            string file = Scraper.getFiles(path)[i];

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

                string newFolder = "week_" + weekCount + "_day_" + dayCount;
                string folderPath = Scraper.folders[dayCount].ToString();
                Dictionary<string, List<string>> contentRaw = new Dictionary<string, List<string>>();

                //Directory.CreateDirectory(savePath + newFolder);

                Console.WriteLine("opening file " + file);
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


