using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace Excel2Json
{
    class Import
    {
        public void Read(string filename)
        {
            string str;
            List<string> content = new List<string>(); 
            int rw = 0;
            int cl = 0;
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(filename);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];

            Excel.Range range = ws.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;
            
            for (int rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (int cCnt = 1; cCnt <= cl; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    content.Add(str);
                    Console.Write(str);
                }
            }

            //foreach (string i in content)
            //{
            //    Console.Write(i);
            //}

            // Get worksheet names
            //foreach (Excel.Worksheet sh in wb.Worksheets)
            //    Console.WriteLine(sh.Name);

            wb.Close();
            excel.Quit();
        }
    }


}

