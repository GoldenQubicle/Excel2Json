using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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



        public void Horizontal(Dictionary<string, List<string>> content)
        {
            // should happen outside here - as lists are needed for each direction. . =) 
            List<string> words = content["Words"];
            List<string> letters = content["Letters"];


            //foreach (string word in words)
            string word = "kippen";
            {
                foreach (char c in word)
                {
                    string letter = c.ToString();
                    if (letters.Contains(letter))
                    {
                        int index = letters.IndexOf(letter);
                        int remain = index % cols;

                        int checkLength = index + word.Count();
                        Console.WriteLine(remain);
                        break;

                    }
                }
            }

        }
        /*
         for this to work I need to have a letter & word array + col&row array in the first place
         what needs to happen here is 
         - iterate over word list
         - take first letter
         - look up first letter in letter list
         - check letterPos + wordLength < dimension 
         - if yes, regex match?
            - if yes, write out letterPos = startCol, startRow & letterPos + wordLengh (AND DIRECTION!) = endCol, endRow

         
         */

        public void getCol(Dictionary<string, List<string>> content)
        {
            List<string> col = content.ElementAt(0).Value;

            foreach (var i in col)
            {
                int.TryParse(i, out cols);
                //Console.WriteLine("cols = " + cols);
            }

        }

        public void getRow(Dictionary<string, List<string>> content)
        {
            List<string> row = content.ElementAt(1).Value;

            foreach (var i in row)
            {
                int.TryParse(i, out rows);
                //Console.WriteLine("rows = " + rows);
            }

        }
    }
}
