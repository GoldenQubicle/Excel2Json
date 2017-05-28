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
    class JSONHelper
    {
        public Dictionary<string, int> getRowsColumns()
        {
            /*
            !!! hardcoded dimensions !!!
            is temporary, will have to be determined on filenames, i.e. level, later on
            and pass level along so to set the correct col & row count 
            */
            int col = 12;
            int row = 12;

            Dictionary<string, int> rowcolumn = new Dictionary<string, int>();
            List<string> cols = new List<string>();
            List<string> rows = new List<string>();

            cols.Add(col.ToString());
            rows.Add(row.ToString());
            rowcolumn.Add("columns", col);
            rowcolumn.Add("rows", row);

            return rowcolumn;
        }

        public List<string> loadRaw()
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
