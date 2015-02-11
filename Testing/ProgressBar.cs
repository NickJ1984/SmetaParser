using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    class ProgressBar
    {
        private float percent;
        private int max;
        private int current;
        private bool isFirstOutput;
        private int top;
        private int left;
        public bool percentOutput { set; get; }

        //public void ProgressBar() { }
        public ProgressBar(int maxValue) { max = maxValue; }

        private void perOutput()
        {
            Console.SetCursorPosition(left, top);
            Console.WriteLine("Progress: {0:0}%", percent);
        }

        private void valOutput()
        {
            Console.SetCursorPosition(left, top);
            Console.WriteLine("Progress: {0}/{1}", max, current);
        }

        public void NextStep(int Val = -1)
        {
            if (Val < 2) current++;
            else current += Val;

            if (current <= max) percent = ((float)current / max) * 100;
        }

        public void Output(int pTop = -1, int pLeft = -1)
        {
            if (isFirstOutput)
            {
                top = Console.CursorTop;
                left = Console.CursorLeft;
                isFirstOutput = true;
            }
            if (pTop > 0) top = pTop;
            if (pLeft > 0) left = pLeft;

            if (percentOutput) perOutput();
            else valOutput();

        }

    }
}
