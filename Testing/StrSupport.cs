using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    static class suppStr    
    {
        static public int[] symbPosSrch(string text, char symb, int startPos = 0, int finPos = 0)
        {
            int[] arr = new int[1];
            int tmp = 0;
            int count = 0;

            if(text == "") return null;

            if (startPos > 0 || finPos > 0)
            {
                if(finPos <= startPos) finPos = text.Length;

                text = text.Substring(startPos, finPos);
            }


            do
            {
                tmp = text.IndexOf(symb, tmp);
                if (tmp >= 0)
                {
                    Array.Resize<int>(ref arr, ++count);
                    arr[count - 1] = tmp++;
                }
            } while (tmp >= 0);

            if (count > 0) return arr;
            else return null;
        }

        static public int symbSum(string text, char sym)
        {
            int result = 0;
            int tmp = 0;

            if(text == "" || sym == null) return 0;
            do
            {
                tmp = text.IndexOf(sym, tmp);
                if (tmp >= 0) { result++; tmp++; }

            } while (tmp >= 0);
            return result;
        }

    }
}
