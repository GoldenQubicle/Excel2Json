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
    class JSONBuilder
    {
        /*
        !!! hardcoded dimensions !!!
        is temporary, will have to be determined on filenames, i.e. level, later on
        and pass level along so to set the correct col & row count 
        */
        int col = 11;
        int row = 11;

        Dictionary<string, List<string>> content = new Dictionary<string, List<string>>();
        List<string> cols = new List<string>();
        List<string> rows = new List<string>();

        public Dictionary<string, List<string>> Build()
        {
            cols.Add(col.ToString());
            rows.Add(row.ToString());
            //content.Add("Columns", cols);
            //content.Add("Rows", rows);
      
            Console.WriteLine("build json");
            return content;

        }

       

        public List<string> LoadRaw()
        {
            List<string> contentRaw = new List<string>();
            using (StreamReader r = new StreamReader("contentRaw.json"))
            {
                string json = r.ReadToEnd();
                List<string> items = JsonConvert.DeserializeObject<List<string>>(json);
                contentRaw = items;
            }
            Console.WriteLine("read in contentRaw.json");
            return contentRaw;
        }
    }


}
