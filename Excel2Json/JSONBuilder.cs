using System;
using System.Collections.Generic;
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
        List<string> letters = new List<string>();
        List<string> words = new List<string>();
        List<string> solution = new List<string>();
        List<string> info = new List<string>();


        string sol = "Oplossing";

        public Dictionary<string, List<string>> Build()
        {
            cols.Add(col.ToString());
            rows.Add(row.ToString());
            content.Add("Columns", cols);
            content.Add("Rows", rows);
            content.Add("Letters", letters);
            content.Add("Words", words);
            content.Add("Solution", solution);
            content.Add("Info", info);

            return content;

        }

        public void SplitScrubbedList(List<string> scrubbedContent)
        {
            foreach (string i in scrubbedContent)
            {
                int count = 0;
                if (Regex.IsMatch(i, sol, RegexOptions.IgnoreCase))
                {
                    solution.Add(i);
                    int uitleg = scrubbedContent.IndexOf(i) + 1;
                    info.Add(scrubbedContent[uitleg]);
                    break;
                }
                foreach (char c in i)
                {
                    if (!char.IsWhiteSpace(c))
                    {
                        count++;
                    }
                }
                if (count == 1)
                {
                    letters.Add(i);
                }
                else
                {
                    words.Add(i);
                }
            }

        }
    }
}
