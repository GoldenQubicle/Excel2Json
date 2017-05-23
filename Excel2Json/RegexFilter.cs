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
    class RegexFilter
    {
        //List<string> contentRaw = new List<string>();
        List<string> content = new List<string>();
        string[] toIgnore = { "Spel", "kleurcode" };

        // obviously need to disable LoadJson when parsing excel files =) 
        public void LoadJson()
        {
            using (StreamReader r = new StreamReader("contentRaw.json"))
            {
                string json = r.ReadToEnd();
                List<string> items = JsonConvert.DeserializeObject<List<string>>(json);
                //contentRaw = items;
            }
        }

        
        public List<string> Filter(List<string> contentRaw)
        {
            List<int> ignoreIndex = new List<int>();
            List<string> scrubbed = new List<string>();

            // eliminate duplicates
            int cut = 0;
            foreach (string i in contentRaw)
            {
                if(Regex.IsMatch(i, toIgnore[0], RegexOptions.IgnoreCase))
                {
                   cut = contentRaw.IndexOf(i);
                }
            }
            for(var i = cut; i < contentRaw.Count; i++)
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
                    content.Add(keep);
                }
            }

            Console.WriteLine("regex filtering done");
            return content;

            //// save scrubbed content, temporary
            //string scrubbedContent = JsonConvert.SerializeObject(content);
            //File.WriteAllText("scrubbedContent.json", scrubbedContent);
        }
    }
}
