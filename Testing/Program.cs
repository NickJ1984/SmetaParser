 using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApplication1
{
    class Program
    {
      
        static void Main(string[] args)
        {
            
            ExcelIO eio = new ExcelIO(@"C:\WORK\Log_2015-01-16_15-06-27.xlsx");
            LogStructure ls = new LogStructure(eio);
            eio.Open();
            //eio.find_once("1234", "A1", "A3");
            ls.buildStructure();
            eio.Quit();
            Console.ReadLine();
        }
    }
}
