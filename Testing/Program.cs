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
            //Log_2015-01-12_15-51-37.xlsx
            //Log_2014-12-30_15-25-53.xlsx
            //ExcelIO eio = new ExcelIO(@"D:\WORK\Урбан-Груп\Программирование\TestingSmetaParser\SmetaParser\Archive_logs\Log_2014-12-30_15-25-53.xlsx");
            ExcelIO eio = new ExcelIO(@"C:\WORK\Log_2015-01-16_15-06-27.xlsx");
            structureBuilder ls = new structureBuilder(eio);
            eio.Open();
            //eio.find_once("1234", "A1", "A3");
            //ls.buildStructure();
            string str = eio.getCellValue(9, 15);
            eio.Quit();
            Console.ReadLine();
        }
    }
}
