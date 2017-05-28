using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System.IO;

namespace Excel2Json
{
    class Import
    {
        Excel.Application excel = new Excel.Application();

       List<string> contentRaw = new List<string>();

        
        // Excel.UsedRange to read a single sheet from workbook, be wary null values. . 
        public List<string> Read(string filename)
        {
            string str;
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
                    if (str != null)
                    {
                        contentRaw.Add(str);
                    }
                }
            }

            //foreach (string i in contentRaw)
            //{
            //    Console.Write(i);
            //}

            wb.Close();
            excel.Quit();
            Console.WriteLine("read excel sheet");
            return contentRaw;
        }

        
    }
    }

