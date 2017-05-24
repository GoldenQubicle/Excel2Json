﻿using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Excel2Json
{
    class Parser
    {
        List<string> contentScrubbed = new List<string>();
        List<string> letters = new List<string>();
        List<string> words = new List<string>();
        List<string> solution = new List<string>();
        List<string> info = new List<string>();

        string sol = "Oplossing";
        string[] toIgnore = { "Spel", "kleurcode" };

        public List<string> scrubContent(List<string> contentRaw)
        {

            List<int> ignoreIndex = new List<int>();
            List<string> scrubbed = new List<string>();

            // eliminate duplicates
            int cut = 0;
            foreach (string i in contentRaw)
            {
                if (Regex.IsMatch(i, toIgnore[0], RegexOptions.IgnoreCase))
                {
                    cut = contentRaw.IndexOf(i);
                }
            }
            for (var i = cut; i < contentRaw.Count; i++)
            {
                scrubbed.Add(contentRaw[i]);
            }

            // get index ignore items
            foreach (string i in toIgnore)
            {
                foreach (string j in scrubbed)
                {
                    if (Regex.IsMatch(j, i, RegexOptions.IgnoreCase))
                    {
                        int pos = scrubbed.IndexOf(j);
                        ignoreIndex.Add(pos);
                    }

                }
            }

            // filter remainder on ignoreIndex
            foreach (string keep in scrubbed)
            {
                if (!ignoreIndex.Contains(scrubbed.IndexOf(keep)))
                {
                    contentScrubbed.Add(keep);
                }
            }

            Console.WriteLine("scrubbed data");
            return contentScrubbed;

            //// save scrubbed content, temporary
            //string scrubbedContent = JsonConvert.SerializeObject(content);
            //File.WriteAllText("scrubbedContent.json", scrubbedContent);
        }

        public Dictionary<string, List<string>> getSolutionInfo(List<string> contentToBeSplit)
        {
            Dictionary<string, List<string>> contentSplit = new Dictionary<string, List<string>>();

            string[] separators = new string[] { " ", ":" };

            foreach (string i in contentToBeSplit)
            {
                if (Regex.IsMatch(i, sol, RegexOptions.IgnoreCase))
                {
                    foreach (string word in i.Split(separators, StringSplitOptions.RemoveEmptyEntries))
                    {
                        solution.Add(word);
                    }
                    int uitleg = contentToBeSplit.IndexOf(i) + 1;
                    info.Add(contentToBeSplit[uitleg]);
                }

            }
            solution.RemoveAt(0);
            contentSplit.Add("solution", solution);
            contentSplit.Add("info", info);

            return contentSplit;
        }

        public Dictionary<string, List<string>> getWordsLetters(List<string> contentToBeSplit)
        {
            Dictionary<string, List<string>> contentSplit = new Dictionary<string, List<string>>();

            foreach (string i in contentToBeSplit)
            {
                int count = 0;
                if (Regex.IsMatch(i, sol, RegexOptions.IgnoreCase))
                {
                    break;
                }
                foreach (char c in i)
                {
                    if (!char.IsWhiteSpace(c))
                    {
                        count++;
                    }
                }
                if (count == 1 || count == 2)
                {
                    letters.Add(i);
                }
                else
                {
                    words.Add(i);
                }
            }
            contentSplit.Add("letters", letters);
            contentSplit.Add("words", words);
   
            Console.WriteLine("split filtered list");
            return contentSplit;
        }
    }
}