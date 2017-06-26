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
        public static string sol = "Oplossing:";
        public static string[] toIgnore = { "kleurcode" };

        public static List<string> scrubContent(List<string> contentRaw)
        {
            List<string> contentScrubbed = new List<string>();
            List<int> ignoreIndex = new List<int>();
            List<string> scrubbed = new List<string>();

            // yeah this is bit useless now because of alternative split 'n scrub
            scrubbed = contentRaw;

            // get index ignore items     
            foreach (string j in scrubbed)
            {
                if (Regex.IsMatch(j, "kleurcode", RegexOptions.IgnoreCase))
                {
                    int pos = scrubbed.IndexOf(j);
                    ignoreIndex.Add(pos);
                }

            }

            // filter  on ignoreIndex
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
        }

        public static Dictionary<string, List<string>> getSolutionInfo(List<string> contentToBeSplit)
        {
            Dictionary<string, List<string>> contentSplit = new Dictionary<string, List<string>>();
            List<string> solution = new List<string>();
            List<string> info = new List<string>();

            string[] separators = new string[] { " ", ":", ".", "," };
            string getSol;

            for (int i = contentToBeSplit.Count; i > 0; i--)
            {
                getSol = contentToBeSplit[i - 1];

                if (Regex.IsMatch(getSol, sol, RegexOptions.IgnoreCase))
                {

                    foreach (string word in getSol.Split(separators, StringSplitOptions.RemoveEmptyEntries))
                    {

                        if (word.Contains("solution")) // remove the solution clue added at import
                        {
                            string lastWord = word.Remove(word.Length - 8);
                            solution.Add(lastWord);
                        }
                        else
                        {
                            solution.Add(word);
                        }
                    }
                    int uitleg = contentToBeSplit.IndexOf(getSol) + 1;
                    info.Add(contentToBeSplit[uitleg]);
                }

            }

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
            Console.WriteLine("getWordsLetters");
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
                // aaarggghhh wordlijst can contain 'ei' as word. . . 

                {
                    if (i == "ei")
                    {
                        words.Add(i);
                    } else
                    {
                        letters.Add(i);
                    }
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
