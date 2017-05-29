using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Excel2Json
{
    class CalcWordPos
    {
        int cols;
        int rows;

        public Dictionary<string, List<int>> FindFirstLetter(Dictionary<string, List<string>> content)
        {
            // okidoki so accidentally made horizontal search here as well
            // obv needs to be factored out ~ for now check if json can be loaded client side
            // also somehow this returns one more than the # of words in words. . weird 
            // doesnt really matter since it needs refactoring anyway
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
                        endCol.Add((index % cols) + word.Count());
                        endRow.Add(index / rows);
                    }

                    if (VerticalSearch(index, word, letters))
                    {
                        startCol.Add(index % cols);
                        startRow.Add(index / rows);
                        endCol.Add(index % cols);
                        endRow.Add((index / rows) + word.Count());
                    }

                    index++;
                }
            }

            ColRow.Add("startCol", startCol);
            ColRow.Add("startRow", startRow);
            ColRow.Add("endCol", endCol);
            ColRow.Add("endRow", endRow);

            return ColRow;
        }


        public bool VerticalSearch(int index, string word, List<string> letters)
        {
            if (((index / rows) + word.Count()) <= rows)
            {
                string letterWord = "";

                for (int i = index; i < (index + (word.Count() * rows)); i += rows)
                {
                    letterWord += letters[i];
                }

                if (letterWord == word)
                {
                    return true;
                }
            }

            return false;
        }


        public bool HorizontalSearch(int index, string word, List<string> letters)
        {
            if (((index % cols) + word.Count()) <= cols)
            {
                string letterWord = "";

                for (int i = index; i < (index + word.Count()); i++)
                {
                    letterWord += letters[i];
                }

                if (letterWord == word)
                {
                    return true;
                }
            }

            return false;
        }




        public Dictionary<string, int> getColRow(string lvl)
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

        // obsolete, was used in cojunction with json helper
        //public void getColRow(Dictionary<string, int> colrow)
        //{
        //     cols = colrow["columns"];
        //     rows = colrow["rows"];
        //}

    }
}
