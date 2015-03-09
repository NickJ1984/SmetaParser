using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{

    #region Structures

    [Serializable]
    public struct db_smetaRecord
    {
        public string Code;
        public string Name;
        public string Object;
        public string Project;
        public DateTime DateLoad;
        public bool Loaded;
    }

    [Serializable]
    public struct db_errorRecord
    {
        public int Number;
        public string Description;
        //public int RecordIndex;
    }
    
    [Serializable]
    public struct db_logfileRecord
    {
        public ust_LogFileDescription File;
        public ust_LogSmetaDescription SmetaDescription;
        public List<ust_LogSmetaData> Data;
        public List<db_errorRecord> Errors;
    }

    [Serializable]
    public struct db_record
    {
        public db_smetaRecord Smeta;
        public List<db_logfileRecord> Logs;
        public int Index;
    }

    #endregion

    [Serializable]
    class DBShell
    {
        public List<db_record> DB { get; private set; }
        private int CurrentIndex = -1;
        private List<ust_LogFileDescription> Logs;

        public DBShell() 
        {
            DB = new List<db_record>();
            Logs = new List<ust_LogFileDescription>();
        }

        #region Add methods

        private void AddElement()
        {
            if (DB == null) DB = new List<db_record>();

            db_record rec = new db_record();
            rec.Logs = new List<db_logfileRecord>();
            ++CurrentIndex;
            rec.Index = CurrentIndex;
            DB.Add(rec);
        }

        private void AddSmetaInfo(int index, ust_LogSmetaDescription lsd)
        {
            db_record dbr = DB.ElementAt(index);
            db_smetaRecord sr = new db_smetaRecord();
            sr.Code = lsd.Smeta.Code;
            
            sr.DateLoad = lsd.LoadTime;
            sr.Name = lsd.Smeta.Name;
            sr.Object = lsd.Smeta.Object;
            sr.Project = lsd.Smeta.Project;
            sr.Loaded = lsd.Loaded;
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
                    { return lsd.Event.Contains("Ошибка"); });
           db_errorRecord eR = new db_errorRecord();
           foreach(ust_LogSmetaData err in errors)
           {
               eR.Description = null; eR.Number = 0; //reinit
               int iOfDot = err.Event.IndexOf('.');
               int iOfSpc = err.Event.IndexOf(' ');

               if (iOfDot > 0 && iOfSpc > 0)
               {
                   eR.Number = int.Parse(err.Event.Substring(iOfSpc + 1, iOfDot - iOfSpc - 1));
                   eR.Description = err.Event.Substring(iOfDot + 2);
               }
               else
               {
                   eR.Number = 0;
                   eR.Description = err.Event;
               }

               lfr.Errors.Add(eR);
           }

        }
        #endregion
        private void AddLogFileRecord(int index, ust_LogFile lf)
        {
            db_logfileRecord lfr = new db_logfileRecord();
            lfr.Errors = new List<db_errorRecord>();
            lfr.Data = new List<ust_LogSmetaData>();

            string Code = DB[index].Smeta.Code;
            int ind = alfr_findSmetaIndex(Code, lf);

            lfr.File = lf.File;
            lfr.Data.AddRange(lf.Body[ind].Data);
            lfr.SmetaDescription = lf.Body[ind].Description;
            alfr_fillErrors(ref lfr);
            DB[index].Logs.Add(lfr);
        }

        private void AddUstLogSmeta(int DBindex, int BodyIndex, ust_LogFile lf)
        {
            db_logfileRecord lfr = new db_logfileRecord();
            lfr.Errors = new List<db_errorRecord>();
            lfr.Data = new List<ust_LogSmetaData>();

            lfr.File = lf.File;
            if (lf.Body[BodyIndex].Data != null) lfr.Data = lf.Body[BodyIndex].Data.ToList();
            lfr.SmetaDescription = lf.Body[BodyIndex].Description;
            alfr_fillErrors(ref lfr);
            DB[DBindex].Logs.Add(lfr);
        }

        public void AddUstLogFile(ust_LogFile lf)
        {
            ust_LogFileDescription lfd = lf.File;
            int index = -1;

            int LFIndex = Logs.FindIndex((ust_LogFileDescription ulfd) => ulfd.FileName == lf.File.FileName);
            if (LFIndex < 0) Logs.Add(lf.File);
            else return; //если в базе уже есть подобный файл, значит завершаем обработку

            for(int i = 0; i <= lf.Body.Count() - 1; i++)
            {
                index = FindEqSmeta(lf.Body[i].Description.Smeta.Code);
                if (index < 0)
                {
                    AddElement();
                    index = CurrentIndex;
                    AddSmetaInfo(index, lf.Body[i].Description);
                }
                AddUstLogSmeta(index, i, lf);
            }
        }

        

        #endregion

        #region Search methods
        private int FindEqSmeta(string Code)
        {
            return DB.FindIndex((db_record dbr) => Code == dbr.Smeta.Code);
        }
        #endregion

        #region Sort methods

        private void SortLogFiles()
        {
            Logs.Sort(delegate(ust_LogFileDescription lfd, ust_LogFileDescription lfd2)
            { return lfd.DateOfCreation.CompareTo(lfd2.DateOfCreation); });
        }

        public void ActualizeDB()
        {
            SortLogFiles();
            for (int i = 0; i < DB.Count; i++)
            {
                DB[i].Logs.Sort(delegate(db_logfileRecord x, db_logfileRecord y)
                { return x.File.DateOfCreation.CompareTo(y.File.DateOfCreation); });
                
                AddSmetaInfo(i, DB[i].Logs[DB[i].Logs.Count - 1].SmetaDescription);
            }
        }

        #endregion


    }

}
