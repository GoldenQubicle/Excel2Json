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
        Excel.Application excel = new Excel.Application();

        List<string> letters = new List<string>();
        List<string> words = new List<string>();
        List<string> contentRaw = new List<string>();

        public List<string> WordArray(string filename)
        {
            Excel.Workbook wb = excel.Workbooks.Open(filename);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];

            int startCol = 16;
            int startRow = 5;
            int sizeCol = 2;
            int sizeRow = 7;
            string word;


            for (int cols = startCol; cols <= (startCol + sizeCol); cols++)
            {
                for (int rows = startRow; rows <= (startRow + sizeRow); rows++)
                {
                    word = (string)(ws.Cells[rows, cols] as Excel.Range).Value2;
                    if (word != null)
                    {
                        words.Add(word);
                    }
                }
            }
            return words;
        }

        public List<string> LetterArray(string filename)
        {
            Excel.Workbook wb = excel.Workbooks.Open(filename);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];

            int startCol = 3;
            int startRow = 5;
            int sizeCol = 11;
            int sizeRow = 11;
            string letter;

            for (int cols = startCol; cols <= (startCol + sizeCol); cols++)
            {
                for (int rows = startRow; rows <= (startRow + sizeRow); rows++)
                {
                    letter = (string)(ws.Cells[rows, cols] as Excel.Range).Value2;
                    if (letter != null)
                    {
                        letters.Add(letter);
                    }
                }
            }

            //Dictionary<string, List<string>> letterArray = new Dictionary<string, List<string>>
            //{
            //    {"letters", letters}
            //};

            return letters;
        }

        // Excel.UsedRange to read single sheet from workbook, be wary null values. . 
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

