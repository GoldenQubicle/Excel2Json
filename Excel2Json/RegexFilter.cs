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
        List<string> content = new List<string>();
        string header = "Spel";


        public void LoadJson()
        {
            using (StreamReader r = new StreamReader("contentRaw.json"))
            {
                string json = r.ReadToEnd();
                List<string> items = JsonConvert.DeserializeObject<List<string>>(json);
                content = items;
                //foreach (var i in items)
                //{
                //    Console.Write(i);
                //}
            }
        }

        public void Filter()
        {
            List<int>toDelete = new List<int>();

            foreach (string i in content)
            {
                if (Regex.IsMatch(i, header, RegexOptions.IgnoreCase))
                {
                    int pos = content.IndexOf(i);
                    toDelete.Add(pos);
                    Console.WriteLine("check I guess" + pos.ToString() + toDelete.Count());
                }
            }
        }

    }
}
