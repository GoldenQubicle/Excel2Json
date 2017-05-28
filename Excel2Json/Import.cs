using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System.IO;
using System.Runtime.InteropServices;

namespace Excel2Json
{
    class Import
    {
           
        public Dictionary<string, List<string>> readFile(string filename)
        {
            Dictionary<string, List<string>> singleXLSX = new Dictionary<string, List<string>>();
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(filename);

            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                Console.WriteLine(filename);
                //determineLevels(ws.Name);
                singleXLSX.Add(determineLevels(ws.Name), SingleSheet(ws));
            }

            wb.Close();
            Marshal.ReleaseComObject(wb);

            excel.Quit();
            Marshal.ReleaseComObject(excel);


            return singleXLSX;
        }

        public string determineLevels(string wsName)
        {
            Dictionary<string, string> levelSublevel = new Dictionary<string, string>();

            string[] levels = { "makkelijk ", "middel ", "moelijk ", "moeilijk + " };

            string[] sublevels = { "(makkelijk)", "(middel)", "(moeilijk)" };

            foreach (string level in levels)
            {
                foreach (string sublevel in sublevels)
                {
                    levelSublevel.Add(level + sublevel, Array.FindIndex(levels, row => row.Contains(level)).ToString() + Array.FindIndex(sublevels, row => row.Contains(sublevel)).ToString());
                }

            }

            string lvl =  levelSublevel[wsName];
            return lvl; 

            //foreach (KeyValuePair<string, string> kvp in levelSublevel)
            //{
            //    Console.WriteLine("Key = {0}, Value = {1}", kvp.Key, kvp.Value);
            //}

        }

        public List<string> SingleSheet(Excel.Worksheet ws)
        {
            List<string> contentRaw = new List<string>();

            string str;
            int rw = 0;
            int cl = 0;

            Excel.Range range = ws.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            for (int rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (int cCnt = 1; cCnt <= cl; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    if (str != null)
                    {
                        contentRaw.Add(str);
                    }
                }
            }

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(ws);

            Console.WriteLine("read excel sheet");
            return contentRaw;
        }
    }
}

