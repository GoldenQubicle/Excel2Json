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

        public void Save(List<string> content)
        {

            string dataRaw = JsonConvert.SerializeObject(content);
            File.WriteAllText("dataRaw.json", dataRaw);

        }
    }
}
