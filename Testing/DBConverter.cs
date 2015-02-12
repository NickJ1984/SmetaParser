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
        public int RecordIndex;
    }

    public struct db_logfileRecord
    {
        public ust_LogFileDescription File;
        public ust_LogSmetaDescription SmetaDescription;
        public List<ust_LogSmetaData> Data;
        public List<db_errorRecord> Errors;

        public string smetaError;
    }

    public struct db_record
    {
        public db_smetaRecord Smeta;
        public List<db_logfileRecord> Logs;
        public int Index;
    }


    class DB
    {
        public List<db_record> DB { get; private set; }

        public DB() { }

        public void NewRecord(db_smetaRecord smeta)
        {
            if (DB == null) DB = new List<db_record>();

            db_record rec = new db_record();
            rec.Smeta = smeta;
            DB.Add(rec);

            //finish = false;
        }

        public void AddLogFileRecord(ust_LogFile logfile, int LFindex = 0, int DBindex = -1)
        {
            db_logfileRecord lfR = new db_logfileRecord();
            lfR.File = logfile.File;
            lfR.Data = logfile.Body[LFindex].Data.ToList<ust_LogSmetaData>();
            lfR.SmetaDescription = logfile.Body[LFindex].Description;
            DB[DBindex].Logs.Add(lfR);

                
        }

        private void errorRecFill(int DBIndex)
        {
            List<ust_LogSmetaData> Data = new List<ust_LogSmetaData>(DB[DBIndex].Logs[DB[DBIndex].Logs.Count - 1].Data);
            string mrk = "Ошибка";



        }

        private bool isRange(int index)
        {
            if(DB == null) return false;

            return (index >= 0 && index <= (DB.Count - 1)) ? true : false;
        }

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
