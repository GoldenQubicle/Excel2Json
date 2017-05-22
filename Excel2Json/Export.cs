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

        public void Save(Dictionary<string, List<string>> content)
        {

            string dataRaw = JsonConvert.SerializeObject(content);
            File.WriteAllText("dataRaw.json", dataRaw);
            Console.WriteLine("saved file, closing now. Bye!");
        }
    }
}
