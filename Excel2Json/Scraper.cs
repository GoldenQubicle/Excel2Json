using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Json
{
    class Scraper
    {

        public void RetrieveFileName(string path)
        {
            /*
             needs to do recursive folder search look for 'WOORDZOEKER' folder

            issue is: the final json file has themeDay, level & sublevel 
            and themeDay is NOT apperent from the xlsx filename, rather needs to be retrieved from folder name
            what's more: the current folder structure is per week  , i.e. week2, dag1 needs to be translated into themeDay
            */

            foreach (string folder in Directory.GetDirectories(path))
            {
                if (folder.Contains("WOORDZOEKER"))
                {
                    foreach (string file in Directory.GetFiles(folder))
                    {
                        Console.WriteLine(file);
                    }
                }
                RetrieveFileName(folder);
            }
        }
    }
}
