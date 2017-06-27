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
    public static class Import
    {
        public static int index;

        public static Dictionary<string, List<string>> readFile(string filename, Dictionary<string, string> lvl)
        {
            Dictionary<string, List<string>> singleXLSX = new Dictionary<string, List<string>>();
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(filename);

            // single sheet debug dev
            //Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[3];
            //Console.WriteLine("processing sheet: " + ws.Name);
            //checkBorders(ws);
            //singleXLSX.Add(determineLevels(ws.Name, lvl), SingleSheet(ws, index));

            // get all sheets from workbook
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                Console.WriteLine("processing sheet: " + ws.Name);
                checkBorders(ws);
                singleXLSX.Add(determineLevels(ws.Name, lvl), SingleSheet(ws, index));
            }

            wb.Close();
            Marshal.ReleaseComObject(wb);

            excel.Quit();
            Marshal.ReleaseComObject(excel);

            return singleXLSX;
        }

        public static void checkBorders(Excel.Worksheet ws)
        {
            Excel.Range range = ws.UsedRange;

            int rw = range.Rows.Count;
            int cl = range.Columns.Count;

            for (int rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (int cCnt = 1; cCnt <= cl; cCnt++)
                {
                    Excel.Range checkBorder = range.Cells[rCnt, cCnt] as Excel.Range;

                    Excel.Borders border = checkBorder.Borders;

                   if (border.LineStyle != Excel.XlLineStyle.xlLineStyleNone.GetHashCode() && rCnt > 20)
                    {
                        index = rCnt;                   
                        return;
                    }
                }
            }
        }
        
        public static List<string> SingleSheet(Excel.Worksheet ws, int rowIndex)
        {
            List<string> contentRaw = new List<string>();
            List<int> indexColor = new List<int>();
            string str;
            int rw = 0;
            int cl = 0;

            Excel.Range range = ws.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            for (int rCnt = rowIndex; rCnt <= rw; rCnt++)
            {
                for (int cCnt = 1; cCnt <= cl; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    if (str != null)
                    {
                        //if (str.Contains("ij"))
                        //{
                        //    str = str.Replace("ij", "y");
                        //}

                        if ((range.Cells[rCnt, cCnt] as Excel.Range).Interior.Color == 13421823)
                        {
                            // color pink = long 13421823
                            // add solution clue to the letter, based off background color cells
                            str += "solution";
                        }
                        //Console.WriteLine(str);
                        contentRaw.Add(str);
                    }
                }
            }

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(ws);

            Console.WriteLine("processing done");
            return contentRaw;
        }

        public static void getGrid(Excel.Worksheet ws)
        {
            Excel.Range range = ws.UsedRange;
            //Excel.Range range = ws.Range["C25", "M35"];
            // Create and fill the grid
            var grid = new ValueGrid();
            int r = 0;
            int c = 0;

            // Loop through the grid-values and store them in the ValueGrid
            foreach (Excel.Range row in range.Rows)
            {
                foreach (Excel.Range cell in row.Cells)
                {
                    //Console.WriteLine($"({r}, {c}): {cell.Value2}");
                    string str = cell.Value2;
                    if (str != null && str.Count() <= 1)
                    {
                        grid.Add(r, c, cell.Value2);
                    }
                    c++; // Next column
                }
                r++; // Next row
                c = 0; // Reset column
            }


            Console.WriteLine(grid.ColumnCount.ToString() + " " + grid.RowCount);

            //Print the found grid
            for (r = 0; r < grid.RowCount; r++)
            {
                // Get all the values for this row
                var values = new List<dynamic>();

                for (c = 0; c < grid.ColumnCount; c++)
                {
                    values.Add(grid[r, c]);
                }
                var print = values
                    .Select(v => (v == null ? "" : v.ToString()).PadLeft(10) + " ") // Give the value some space in the string
                    .ToArray();

                // Print the row
                Console.WriteLine(string.Join("|", print));
            }


            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(ws);

            Console.WriteLine("processing done");
        }

        public static string determineLevels(string wsName, Dictionary<string, string> levelSublevel)
        {
            wsName = wsName.Trim();
            string lvl = levelSublevel[wsName];
            return lvl;
        }

    }
}

