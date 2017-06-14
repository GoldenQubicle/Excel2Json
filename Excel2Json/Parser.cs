using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Excel2Json
{
    public static class Parser
    {
        public static string sol = "Oplossing";
        public static string[] toIgnore = { "Spel", "kleurcode" }; 

        public static List<string> scrubContent(List<string> contentRaw)
        {
            List<string> contentScrubbed = new List<string>();
            List<int> ignoreIndex = new List<int>();
            List<string> scrubbed = new List<string>();

            // eliminate duplicates
            int cut = 0;
            foreach (string i in contentRaw)
            {
                //Console.WriteLine(i);
                if (Regex.IsMatch(i, toIgnore[0], RegexOptions.IgnoreCase))
                {
                    cut = contentRaw.IndexOf(i);
                    //Console.WriteLine(cut);
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
                    //Console.WriteLine(keep);
                }
            }

            Console.WriteLine("scrubbed data");
            return contentScrubbed;

            //// save scrubbed content, temporary
            //string scrubbedContent = JsonConvert.SerializeObject(content);
            //File.WriteAllText("scrubbedContent.json", scrubbedContent);
        }

        public static Dictionary<string, List<string>> getSolutionInfo(List<string> contentToBeSplit)
        {
            Dictionary<string, List<string>> contentSplit = new Dictionary<string, List<string>>();
            List<string> solution = new List<string>();
            List<string> info = new List<string>();

            string[] separators = new string[] { " ", ":", ".", "," };

            foreach (string i in contentToBeSplit)
            {
                if (Regex.IsMatch(i, sol, RegexOptions.IgnoreCase))
                {
                    //Console.WriteLine(i.ToString());

                    foreach (string word in i.Split(separators, StringSplitOptions.RemoveEmptyEntries))
                    {
                        if (word != "solution") // remove the solution clue added at import
                        {
                            solution.Add(word);
                            //Console.WriteLine(word);
                        }
                    }
                    int uitleg = contentToBeSplit.IndexOf(i) + 1;
                    info.Add(contentToBeSplit[uitleg]);
                }

            }                      
           //Console.WriteLine(solution.Count() + " " + info.Count()); // so nothing is added to solution?!
            
            solution.RemoveAt(0);          
            contentSplit.Add("solution", solution);
            contentSplit.Add("info", info);

            return contentSplit;
        }

        public static Dictionary<string, List<string>> getWordsLetters(List<string> contentToBeSplit)
        {
            Dictionary<string, List<string>> contentSplit = new Dictionary<string, List<string>>();
            List<string> letters = new List<string>();
            List<string> words = new List<string>();

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
                if (count == 1 || count == 2 || i.Contains("solution"))
                {
                    letters.Add(i);
                    //Console.WriteLine(i);
                }
                else if (!i.Contains(" "))
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
