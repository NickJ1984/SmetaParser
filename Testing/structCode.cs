using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    class structCode
    {
        private char Separator;
        private string Code = "";
        private string[] Block;
        private int[] sepPos;
        private int bCount;

        public structCode() 
        {
            
        }
        public structCode(string StructCode) : base()
        {
            Separator = '-';
            bCount = 0;
            Code = StructCode;
            blockCalculate();
        }

        private void blockCalculate()
        {
            int tPos = 0;

            do
            {
                tPos = Code.IndexOf(Separator, tPos);
                if( tPos >= 0 ) 
                {
                    bCount++;
                    Array.Resize<int>(ref sepPos, bCount);
                    sepPos[bCount - 1] = tPos++;
                }
            } while (tPos >= 0);
        }

        public void Debug()
        {
            Console.WriteLine("Анализ кода:");
            Console.WriteLine("Количество блоков: {0}", bCount);
            Console.WriteLine("Позиция разделителя:");
            if (bCount > 0)
            {
                for (int i = 0; i < bCount - 1; i++) Console.WriteLine(sepPos[i]);
            }
        }

        private bool verifySyntax(string str)
        {
            bool ans = false;


            return ans;
        }
    }
}
