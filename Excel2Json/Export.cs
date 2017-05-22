using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Json
{
    class Export
    {
        public void SaveRaw(List<string> contentRaw)
        {
            string dataRaw = JsonConvert.SerializeObject(contentRaw);
            File.WriteAllText("contentRaw.json", dataRaw);
            Console.WriteLine("saved file, closing now. Bye!");

        }

        public void Save(Dictionary<string, List<string>> content)
        {

            string data = JsonConvert.SerializeObject(content);
            File.WriteAllText("data.json", data);
            Console.WriteLine("saved file, closing now. Bye!");
        }
    }
}
