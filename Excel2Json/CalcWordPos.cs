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



        public void FindFirstLetter(Dictionary<string, List<string>> content)
        {
            List<string> words = content["Words"];
            List<string> letters = content["Letters"];

            //foreach (string word in words)
            string word = "kippen";
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
                        Console.WriteLine(letterWord);

                    }


                    //Console.WriteLine(index);

                    // Increment the index.
                    index++;
                }

            }
        }


        public void Horizontal(Dictionary<string, List<string>> content)
        {
            // should happen outside here - as lists are needed for each direction. . =) 
            List<string> words = content["Words"];
            List<string> letters = content["Letters"];

            //foreach (string word in words)
            int index;
            string word = "kippen";
            string letter = word[0].ToString();
            string letterWord = "";

            // check for first letter of word, and if the word will fit in that position
            // if yes, construct string from letter array and see of that matches the word

            foreach (string l in letters)
            {
                if (Regex.IsMatch(l, letter) && ((letters.IndexOf(l) % cols) + word.Count()) <= cols)
                {
                    index = letters.IndexOf(l);

                    for (int i = index; i < (index + word.Count()); i++)
                    {
                        letterWord += letters[i];
                    }

                    if (!Regex.IsMatch(word, letterWord))
                    {
                        Console.WriteLine("check pls" + letterWord);
                    }
                }

            }

        }
        //    if (letters.Contains(letter) && ((letters.IndexOf(letter) % cols) + word.Count()) <= cols)
        //    {
        //        index = letters.IndexOf(letter);

        //        for (int i = index; i < (index + word.Count()); i++)
        //        {
        //            letterWord += letters[i];
        //        }

        //        if (!Regex.IsMatch(word, letterWord))
        //        {
        //            Console.WriteLine("check pls" + letterWord);
        //        }
        //    }
        //}




        //Console.WriteLine("check" + letters.IndexOf(letter) % cols + letters.IndexOf(letter) / rows);


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
