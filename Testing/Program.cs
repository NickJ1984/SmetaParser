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
            
            //ExcelIO eio = new ExcelIO(@"C:\WORK\Log_2015-01-16_15-06-27.xlsx");
            LogStructure ls = new LogStructure();
            ls.addSmeta("111");
            ls.addSmeta("111");
            ls.addSmeta("111");
            ls.addEvent("111");
            ls.addEvent("111"); ls.addEvent("111"); ls.addEvent("111");
            Console.ReadLine();
        }
    }
}
