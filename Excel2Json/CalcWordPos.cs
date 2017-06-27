using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Excel2Json
{
    public static class CalcWordPos
    {
        public static int cols;
        public static int rows;
        public static List<string> foundWords;


        public static Dictionary<string, List<int>> FindWords(Dictionary<string, List<string>> content)
        {
            foundWords = new List<string>();
            Dictionary<string, List<int>> ColRow = new Dictionary<string, List<int>>();
            List<int> startCol = new List<int>();
            List<int> startRow = new List<int>();
            List<int> endCol = new List<int>();
            List<int> endRow = new List<int>();
            List<string> words = content["words"];
            List<string> letters = content["letters"];

            foreach (string word in words)
            {

            
                int index = 0;
                string firstLetter = word[0].ToString();
                while ((index = letters.IndexOf(firstLetter, index)) != -1)
                {

                    if (HorizontalSearch(index, word, letters))
                    {
                        startCol.Add(index % cols);
                        startRow.Add(index / rows);
                        endCol.Add(((index % cols) + (word.Count()) - 1));
                        endRow.Add(index / rows);
                        //Console.WriteLine("woord = " + word + " start col & row = " + (index % cols) + " " + (index / rows) + " end col & row " + ((index % cols) + word.Count()) + " " + (index / rows));
                        break;
                    }

                    else if (VerticalSearch(index, word, letters))
                    {
                        startCol.Add(index % cols);
                        startRow.Add(index / rows);
                        endCol.Add(index % cols);
                        endRow.Add(((index / rows) + (word.Count()) - 1));
                        //Console.WriteLine("woord = " + word + " start col & row = " + (index % cols) + " " + (index / rows) + " end col & row " + (index % cols) + " " + ((index / rows) + word.Count()));
                        break;
                    }

                    else if (DiagonalSearch(index, word, letters))
                    {
                        startCol.Add(index % cols);
                        startRow.Add(index / rows);
                        endCol.Add(((index % cols) + (word.Count()) - 1));
                        endRow.Add(((index / rows) + (word.Count()) - 1));
                        //Console.WriteLine("woord = " + word + " start col & row = " + (index % cols) + " " + (index / rows) + " end col & row " + ((index % cols) + word.Count()) + " " + ((index / rows) + word.Count()));
                        break;
                    }

                    index++;
                }
            }

            ColRow.Add("startCol", startCol);
            ColRow.Add("startRow", startRow);
            ColRow.Add("endCol", endCol);
            ColRow.Add("endRow", endRow);
            Console.WriteLine("found all positions words");
            return ColRow;
        }

        public static Dictionary<string, List<int>> FindSolution(Dictionary<string, List<string>> content)
        {

            Dictionary<string, List<int>> SolColRow = new Dictionary<string, List<int>>();
            List<int> solColStart = new List<int>();
            List<int> solRowStart = new List<int>();
            List<int> solColEnd = new List<int>();
            List<int> solRowEnd = new List<int>();

            List<string> solution = content["solution"];
            List<string> letters = content["letters"];

            foreach (string word in solution)
            {
                int index = 0;
                if (word != "")
                {
                    string firstLetter = word[0].ToString().ToLower();
                    firstLetter += "solution";
                    int wordLength = word.Count();

                    while ((index = letters.IndexOf(firstLetter, index)) != -1)
                    {
                        if (((index % cols) + wordLength) <= cols)
                        {
                            string letterWord = "";

                            for (int i = index; i < (index + wordLength); i++)
                            {
                                if (letters[i].Contains("solution")) // extra check for solution clue
                                {
                                    letterWord += letters[i].Remove(letters[i].Length - 8);

                                    if (letterWord == word.ToLower())
                                    {
                                        solColStart.Add(index % cols);
                                        solRowStart.Add(index / rows);
                                        solColEnd.Add((index % cols) + wordLength);
                                        solRowEnd.Add(index / rows);
                                        break;
                                    }
                                }
                            }
                        }

                        index++;
                    }
                }
            }

            // remove the solution clue from letter array
            for (var i = 0; i < letters.Count(); i++)
            {
                if (letters[i].Contains("solution"))
                {
                    //if (letters[i].Contains("ij"))
                    //{
                    //    letters[i] = letters[i].Remove(2, 8);

                    //}
                    //else
                    //{
                    letters[i] = letters[i].Remove(1, 8);
                    //}
                }
            }

            SolColRow.Add("solColStart", solColStart);
            SolColRow.Add("solRowStart", solRowStart);
            SolColRow.Add("solColEnd", solColEnd);
            SolColRow.Add("solRowEnd", solRowEnd);
            Console.WriteLine("found all positions solution");

            return SolColRow;
        }

        public static bool DiagonalSearch(int index, string word, List<string> letters)
        {

            if (((index % cols) + word.Count()) <= cols && ((index / rows) + word.Count()) <= rows)
            {
                string letterWord = "";

                for (int i = index; i < letters.Count(); i += (cols+1))
                {
                    letterWord += letters[i];
                    Console.WriteLine(letterWord);

                    if (letterWord == word.ToLower() && !foundWords.Contains(letterWord))
                    {
                        foundWords.Add(letterWord);
                        return true;
                    }
                }
            }

            return false;
        }

        public static bool VerticalSearch(int index, string word, List<string> letters)
        {
            // bit hacky here to circumvent rouning down integer division
            //float test = (float)index / (float)rows;
            //double round = Math.Round(test, MidpointRounding.ToEven);
            //index = (int)round;

            if (((index / cols) + word.Count()) <= cols)
            {
                string letterWord = "";

                for (int i = index; i < (index + (word.Count() * cols)); i += cols)
                {
                    letterWord += letters[i];
                    Console.WriteLine(letterWord);

                    if (letterWord == word.ToLower() && !foundWords.Contains(letterWord))
                    {
                        foundWords.Add(letterWord);
                        return true;
                    }
                }
            }

            return false;
        }

        public static bool HorizontalSearch(int index, string word, List<string> letters)
        {
            if (((index % cols) + word.Count()) <= cols && index + word.Count() <= letters.Count())
            {
                string letterWord = "";

                for (int i = index; i < (index + word.Count()); i++)
                {
                    letterWord += letters[i];
                    Console.WriteLine(letterWord);

                    if (letterWord == word.ToLower() && !foundWords.Contains(letterWord))
                    {
                        foundWords.Add(letterWord);
                        return true;
                    }
                }
            }

            return false;
        }

        public static Dictionary<string, int> getColRow(string lvl)
        {
            if (lvl == "00" || lvl == "01" || lvl == "02")
            {
                cols = 11;
                rows = 11;
            }

            if (lvl == "10" || lvl == "11" || lvl == "12")
            {
                cols = 12;
                rows = 12;
            }

            if (lvl == "20" || lvl == "21" || lvl == "22")
            {
                cols = 12;
                rows = 13;
            }

            if (lvl == "30" || lvl == "31" || lvl == "32")
            {
                cols = 13;
                rows = 13;
            }


            Dictionary<string, int> rowcolumn = new Dictionary<string, int>();

            rowcolumn.Add("columns", cols);
            rowcolumn.Add("rows", rows);

            return rowcolumn;
        }
    }
}
