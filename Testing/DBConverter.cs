using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    public struct db_smetaRecord
    {
        public string Code;
        public string Name;
        public string Object;
        public string Project;
        public DateTime DateLoad;
    }

    public struct db_errorRecord
    {
        public int Number;
        public string Description;
        //public int RecordIndex;
    }

    public struct db_logfileRecord
    {
        public ust_LogFileDescription File;
        public ust_LogSmetaDescription SmetaDescription;
        public List<ust_LogSmetaData> Data;
        public List<db_errorRecord> Errors;
    }

    public struct db_record
    {
        public db_smetaRecord Smeta;
        public List<db_logfileRecord> Logs;
        public int Index;
    }


    class DBShell
    {
        public List<db_record> DB { get; private set; }
        
        public DBShell() 
        {
            DB = new List<db_record>();
        }

        #region Add methods

        public void AddElement()
        {
            if (DB == null) DB = new List<db_record>();

            db_record rec = new db_record();
            
            rec.Index = DB.Count - 1;
            DB.Add(rec);
        }

        public void AddSmetaInfo(int index, ust_LogSmetaDescription lsd)
        {
            db_record dbr = DB.ElementAt(index);
            db_smetaRecord sr = new db_smetaRecord();
            sr.Code = lsd.Smeta.Code;
            sr.DateLoad = lsd.LoadTime;
            sr.Name = lsd.Smeta.Name;
            sr.Object = lsd.Smeta.Object;
            sr.Project = lsd.Smeta.Project;
            dbr.Smeta = sr;
            DB[index] = dbr;
        }

        #region alfr service

        private int alfr_findSmetaIndex(string code, ust_LogFile lf)
        {
            for (int i = 0; i <= lf.Body.Length; i++)
                if (code.ToUpper() == lf.Body[i].Description.Smeta.Code.ToUpper()) return i;
            return -1; 
        }

        private void alfr_fillErrors(ref db_logfileRecord lfr)
        {
            if (lfr.Data == null) return;
            if (lfr.Data.Count < 1) return;

            List<ust_LogSmetaData> errors = lfr.Data.FindAll(
                delegate(ust_LogSmetaData lsd) 
                    {
                        return lsd.Event.Contains("Ошибка");
                    });
           db_errorRecord eR = new db_errorRecord();
           foreach(ust_LogSmetaData err in errors)
           {
               eR.Description = null; eR.Number = 0; //reinit
               int iOfDot = err.Event.IndexOf('.');
               int iOfSpc = err.Event.IndexOf(' ');

               eR.Number = int.Parse(err.Event.Substring(iOfSpc + 1, iOfDot - iOfSpc - 1));
               eR.Description = err.Event.Substring(iOfDot + 2);

               lfr.Errors.Add(eR);
           }

        }
        #endregion
        public void AddLogFileRecord(int index, ust_LogFile lf)
        {
            db_logfileRecord lfr = new db_logfileRecord();
            string Code = DB[index].Smeta.Code;
            int ind = alfr_findSmetaIndex(Code, lf);

            lfr.File = lf.File;
            lfr.Data.AddRange(lf.Body[ind].Data);
            lfr.SmetaDescription = lf.Body[ind].Description;
            alfr_fillErrors(ref lfr);
            DB[index].Logs.Add(lfr);
        }
        
        #endregion

        #region Search methods

        public int FindEqSmeta(string Code)
        {
            return DB.FindIndex((db_record dbr) => Code == dbr.Smeta.Code);
        }

        #endregion


        /*
        private bool IndexOfError(ust_LogSmetaData lsd)  //предикат метода errorRecFill
        {
            string mrk = "Ошибка";
            return lsd.Event.Contains(mrk);            
        }

        private void errorRecFill(int DBIndex)
        {
            //Ошибка 29. Сумма загружаемого из ГрандСметы акта отличается от суммы соответствующей утвержденной КС-2 в 1С на -72 317,66

            List<ust_LogSmetaData> Data = new List<ust_LogSmetaData>(DB[DBIndex].Logs[DB[DBIndex].Logs.Count - 1].Data);
            int[] eInd = new int[1];
            int cnt = 0;

            int index = Data.FindIndex(0, IndexOfError);
            if (index == -1) return;
            
            while (index != -1)
            {
                Array.Resize<int>(ref eInd, ++cnt);
                eInd[eInd.Length - 1] = index;

                index = Data.FindIndex(index + 1, IndexOfError);
            }

            db_errorRecord er = new db_errorRecord();

            foreach (int ind in eInd)
            {
                er.Description = null; er.Number = 0; er.RecordIndex = 0; //reinit

                string ev = Data[ind].Event;
                int iOfDot = ev.IndexOf('.');
                int iOfSpc = ev.IndexOf(' ');

                er.Number = int.Parse(ev.Substring(iOfSpc + 1, iOfDot - iOfSpc - 1));
                er.RecordIndex = ind;
                er.Description = ev.Substring(iOfDot + 2);

                DB[DBIndex].Logs[DB[DBIndex].Logs.Count - 1].Errors.Add(er);
            }
        }

        private bool isRange(int index)
        {
            if(DB == null) return false;

            return (index >= 0 && index <= (DB.Count - 1)) ? true : false;
        }*/
    }




    class DBConverter
    {
        List<db_record> DB;
        List<string> Indexer;

        public DBConverter() { }


        #region Indexer methods

        public int EqSearch(string Code)
        {
            // -1 == не найдено
            return Indexer.IndexOf(Code);
        }

        public int addIndex(string Code)
        {
            Indexer.Add(Code);
            return Indexer.Count - 1;
        }
        
        #endregion

        #region DB methods

        public void DBNewRecord(ust_LogSmetaDescription lsd)
        {
            

        }


        #endregion

    }
}
