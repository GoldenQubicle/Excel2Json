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

        List<int> startCol = new List<int>();
        List<int> startRow = new List<int>();
        List<int> endCol = new List<int>();
        List<int> endRow = new List<int>();

        public Dictionary<string, List<int>> FindFirstLetter(Dictionary<string, List<string>> content)
        {
            // okidoki so accidentally made horizontal search here as well
            // obv needs to be factored out ~ for now check if json can be loaded client side
            Dictionary<string, List<int>> ColRow = new Dictionary<string, List<int>>();

            List<string> words = content["words"];
            List<string> letters = content["letters"];

            foreach (string word in words)
            {
                int index = 0;
                string firstLetter = word[0].ToString();               
                while ((index = letters.IndexOf(firstLetter, index)) != -1)
                {
                    if(((index % cols) + word.Count()) <= cols){
                        string letterWord = "";

                        for (int i = index; i < (index + word.Count()); i++)
                        {
                            letterWord += letters[i];
                        }

                        if (letterWord == word)
                        {
                            startCol.Add(index % cols);
                            startRow.Add(index / rows);
                            endCol.Add((index % cols) + word.Count());
                            endRow.Add(index / rows);
                        }
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
                                   
        public void getColRow(Dictionary<string, int> colrow)
        {
             cols = colrow["columns"];
             rows = colrow["rows"];
        }

    }
}
