using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Json
{
    public static class Scraper
    {
        public static List<string> folders = new List<string>();

        static List<string> fileNames = new List<string>();

        public static List<string> getFiles(string path)
        {
            foreach (string folder in Directory.GetDirectories(path))
            {
                if (folder.Contains("WOORDZOEKER"))
                {
              
                    foreach (string file in Directory.GetFiles(folder))
                    {
                        fileNames.Add(file);
                        //Console.WriteLine(file);
                    }
                    //Console.WriteLine(folder);
                    folders.Add(folder);
                }
                getFiles(folder);
            }
            return fileNames;
        }

    }
}
