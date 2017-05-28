﻿using System;
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

            determine level on filename
            determine sublevel on sheetname
            scrape folders for filenames  
            also, prolly want to properly close COM object when scraping entire folders . . =) 
            
            */

            //Import import = new Import();
            Scraper scraper = new Scraper();
            Export export = new Export();
            JSONHelper jsonHelper = new JSONHelper();
            Parser parser = new Parser();
            CalcWordPos calcWordPos = new CalcWordPos();
            List<string> content = new List<string>();
            Dictionary<string, List<string>> contentFormattedStrings = new Dictionary<string, List<string>>();
            Dictionary<string, List<int>> contentFormattedInts = new Dictionary<string, List<int>>();
            Dictionary<string, int> contentColRow = new Dictionary<string, int>();

            String filename = "C:\\Users\\Erik\\Desktop\\MAAND 1, WEEK 2\\MAAND 1, WEEK 2, DAG 1\\WOORDZOEKER\\WZ middel LANDBOUW.xlsx";

            scraper.RetrieveFileName(@"C:\Users\Erik\Desktop\MAAND 1, WEEK 2");

            // proper routine needs to go here =) 
            //content = parser.scrubContent(import.Read(filename));

            // temp routine to save single excel sheet as json 
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


