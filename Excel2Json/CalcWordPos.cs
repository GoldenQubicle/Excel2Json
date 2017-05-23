using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Json
{
    class CalcWordPos
    {
        List<int> startCol = new List<int>();
        List<int> startRow = new List<int>();
        List<int> endCol = new List<int>();
        List<int> endRow = new List<int>();

        public void Horizontal()
        {


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
 

    }
}
