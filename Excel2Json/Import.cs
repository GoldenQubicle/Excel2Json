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
           
        public Dictionary<string, List<string>> readFile(string filename, Dictionary<string, string> lvl)
        {
            Dictionary<string, List<string>> singleXLSX = new Dictionary<string, List<string>>();
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(filename);

            // single sheet debug dev
            //Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[3];
            //Console.WriteLine(filename + " sheet " + ws.Name);
            //singleXLSX.Add(determineLevels(ws.Name, lvl), SingleSheet(ws));

            // get all sheets from workbook
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                Console.WriteLine("processing sheet: " + ws.Name);
                singleXLSX.Add(determineLevels(ws.Name, lvl), SingleSheet(ws));
            }

            wb.Close();
            Marshal.ReleaseComObject(wb);

            excel.Quit();
            Marshal.ReleaseComObject(excel);


            return singleXLSX;
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

            Console.WriteLine("processing done");
            return contentRaw;
        }

        public string determineLevels(string wsName, Dictionary<string, string> levelSublevel)
        {
            string lvl = levelSublevel[wsName];
            return lvl;
        }

    }
}

