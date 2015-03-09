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
        static void dbgSort()
        {
            Serializer srl = new Serializer();
            srl.path = @"K:\Git\Test\DBSData.dat";
            long t1 = DateTime.Now.Ticks;
            srl.Read();
            DBShell dbs = (DBShell)srl.obj;
            srl.obj = null;
            System.GC.Collect();
            long t2 = DateTime.Now.Ticks;
            Console.WriteLine("Time to load: {0} ticks", t2 - t1);
            
            Console.WriteLine("Start log files sorting");
            t1 = DateTime.Now.Ticks;
            dbs.SortLogFiles();
            t2 = DateTime.Now.Ticks;
                        
            Console.WriteLine("Finish log files sorting");
            Console.WriteLine("Time: {0} ticks", t2 - t1);
            
            
            Console.WriteLine("Start DB sorting");
            t1 = DateTime.Now.Ticks;
            dbs.ActualizeDB();
            t2 = DateTime.Now.Ticks;
            Console.WriteLine("Finish DB sorting");
            Console.WriteLine("Time: {0} ticks", t2 - t1);
            srl.obj = dbs;
            srl.Write();
            Console.ReadLine();
            
        }

        static void Main(string[] args)
        {
            dbgSort();
            return;

            ust_LogFile[] lf;
            FileIO fio = new FileIO(@"K:\Git\Test\AllLogs");
            Serializer srl = new Serializer();
            srl.path = @"K:\Git\Test\DBSData.dat";

            fio.searchPattern = "*.xlsx";
            fio.scan();
            lf = new ust_LogFile[fio.logfiles.Length];
            ExcelIO eio = new ExcelIO();

            ProgressBar pb = new ProgressBar(lf.Length);
            pb.percentOutput = false;
            
            for (int i = 0; i < lf.Length; i++)
            {
                pb.Information = fio.logfiles[i].FileName;
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

            DBShell dbs = new DBShell();
            foreach (ust_LogFile ulf in lf) dbs.AddUstLogFile(ulf);
            srl.obj = dbs;
            srl.Write();

            Console.ReadLine();
        }
    }
}
