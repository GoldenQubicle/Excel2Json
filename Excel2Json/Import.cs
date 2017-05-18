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
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(filename);

            // Get worksheet names
            foreach (Excel.Worksheet sh in wb.Worksheets)
                Console.WriteLine(sh.Name);



            List<object> letters = new List<object>();

            for (int cols = 3; cols < (cols + 11); cols++)
            {
                for (int rows = 5; rows < (rows + 11); rows++)
                {
                    letters.Add(wb.Sheets["makkelijk (makkelijk)"].Cells[cols, rows].Value2);
                    Console.WriteLine(letters.Count);

                }
            }
            
            foreach (object letter in letters)
            {
                Console.WriteLine(letter);
            }

            wb.Close();
            excel.Quit();
        }
    }


}

