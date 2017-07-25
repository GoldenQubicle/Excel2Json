using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Json
{
    public static class Export
    {

        public static void SaveFinal(Dictionary<string, List<string>> contentString, Dictionary<string, List<int>> contentInt, Dictionary<string, int> colrow, string key, string folderPath)
        {
            Dictionary<string, object> contentFinal = new Dictionary<string, object>();

            contentFinal.Add("columns", colrow["columns"]);
            contentFinal.Add("rows", colrow["rows"]);
            contentFinal.Add("words", contentString["words"].ToList());
            contentFinal.Add("letters", contentString["letters"].ToList());
            contentFinal.Add("solution", contentString["solution"].ToList());
            contentFinal.Add("info", contentString["info"].ToList());
            contentFinal.Add("startCol", contentInt["startCol"].ToList());
            contentFinal.Add("startRow", contentInt["startRow"].ToList());
            contentFinal.Add("endCol", contentInt["endCol"].ToList());
            contentFinal.Add("endRow", contentInt["endRow"].ToList());
            contentFinal.Add("solColStart", contentInt["solColStart"].ToList());
            contentFinal.Add("solRowStart", contentInt["solRowStart"].ToList());
            contentFinal.Add("solColEnd", contentInt["solColEnd"].ToList());
            contentFinal.Add("solRowEnd", contentInt["solRowEnd"].ToList());

            // check to make sure number of letters equals numbers of cells
            if (colrow["rows"] * colrow["columns"] == contentString["letters"].Count()
                && contentInt["startCol"].Count() == contentString["words"].Count())
            {
                string dataFinal = JsonConvert.SerializeObject(contentFinal);
                File.WriteAllText(folderPath + "\\ws-" + key + ".json", dataFinal);
                Console.WriteLine("saved file : " + folderPath + "\\" + key + ".json");
            }
                     else
            {
                Console.WriteLine("welp, not enough letters. . or row / column count is off. Either way, big fat nope =) ");
                Console.WriteLine("columns: " + colrow["columns"]);
                Console.WriteLine("rows: " + colrow["rows"]);
                Console.WriteLine("letters: " + contentString["letters"].Count());
                Console.WriteLine("number of words = " + contentString["words"].Count() + " " + "number of start pos " + contentInt["startCol"].Count());
                string dataFinal = JsonConvert.SerializeObject(contentFinal);
                File.WriteAllText(folderPath + "\\ws-redo" + key + ".json", dataFinal);
                Console.WriteLine("saved file : " + folderPath + "\\" + key + ".json");
            }
        }

        public static void SaveIntermediate(Dictionary<string, List<string>> contentStrings)
        {
            string dataStrings = JsonConvert.SerializeObject(contentStrings);
            //string dataInts = JsonConvert.SerializeObject(contentInts);
            //string data =  String.Join("",dataInts, dataStrings);
            File.WriteAllText("contentTemp.json", dataStrings);
            Console.WriteLine("saved temp file, closing now. Bye!");

        }

        public static void SaveRaw(List<string> contentRaw, string key)
        {
            string dataRaw = JsonConvert.SerializeObject(contentRaw);
            File.WriteAllText("contentRaw" + key + ".json", dataRaw);
            Console.WriteLine("saved file " + key + ", closing now. Bye!");
        }

        public static void Save(Dictionary<string, List<string>> content)
        {

            string data = JsonConvert.SerializeObject(content);
            File.WriteAllText("data.json", data);
            Console.WriteLine("succesfully saved file!");
        }

    }
}
