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
            //ExcelIO eio = new ExcelIO(@"C:\WORK\Log_2015-01-16_15-06-27.xlsx");
            ust_LogFile[] lf;
            FileIO fio = new FileIO(@"C:\WORK\Программирование\C#\Visual Studio\Projects\TestingSmetaParser\Archive_logs");
            fio.searchPattern = "*.xlsx";
            fio.scan();
            lf = new ust_LogFile[fio.logfiles.Length];
            ExcelIO eio = new ExcelIO();

            ProgressBar pb = new ProgressBar(lf.Length);
            pb.percentOutput = true;
            
            for (int i = 0; i < lf.Length; i++)
            {
                pb.NextStep();
                pb.Output();
                lf[i].File = fio.logfiles[i];

                eio.Open(lf[i].File.FullPath);
                
                structureBuilder sb = new structureBuilder(eio);
                sb.buildStructure();
                
                structureReader sr = new structureReader(sb.getData(), eio);
                sr.Read();
                lf[i].Body = sr.smetalog;
                                
                sb = null;
                sr = null;
                eio.CloseWB();
                System.GC.Collect();
            }
            eio.Quit();
            Console.ReadLine();
        }
    }
}
