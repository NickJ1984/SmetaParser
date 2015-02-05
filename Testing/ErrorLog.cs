using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApplication1
{
    #region UserStructures

    public struct ust_LogSmetaRegion
    {
        public string adrSmeta;
        public string adrEvent;
        public int startRow;
        public int eventRow;
        public int endRow;
    }

    public struct ust_LogFile
    {
        public ust_LogFileDescription File;
        public ust_LogSmeta Body;
    }

    public struct ust_LogSmeta
    {
        public ust_LogSmetaDescription Description;
        public ust_LogSmetaData[] Data;
    }

    public struct ust_LogSmetaDescription
    {
        public ust_Smeta Smeta;
        public DateTime LoadTime;
        public bool Loaded;
    }

    public struct ust_Smeta
    {
        public string FileName;
        public string Name;
        public string Code;
        public string Object;
        public string Project;
        public string Number;
        public string Status;
    }

    public struct ust_LogFileDescription
    {
        public string FullPath;
        public string FileName;
        public DateTime DateOfCreation;
    }

    public struct ust_LogSmetaData
    {
        public string ppNumber;
        public string Event;
        public string SysID;
        public string Code1C;
        public string Name;
        public string Description;
    }
    #endregion


    class ErrorLog
    {
        #region Variables
        private const string ext = "*.xlsx";

        private ust_LogFile[] GlobalData;
        private bool isExist;

        public string directory { get; private set; }
        #endregion

        #region Constructors
        public ErrorLog() {}
        
        public ErrorLog(string path) :base()
        {
            changeDirectory(path);
        }
        #endregion

        #region FileIO segment

        private void changeDirectory(string dir)
        {
            if (dir.LastIndexOf('\\') != (dir.Length - 1)) dir = dir + @"\";

            if (Directory.Exists(dir))
            {
                directory = dir;
                isExist = true;
            }
        }

        public void getLogFiles()
        {
            if (!isExist) return;
            int i, j;

            FileIO fio = new FileIO(directory);
            fio.getFiles(ext);

            GlobalData = new ust_LogFile[fio.filePath.Length];

            for (i = 0, j = 0; i < fio.filePath.Length; i++, j++)
            {
            Next:
                if(fio.files[i].IndexOf("~$") >= 0)
                {
                    i++;
                    if(i >= fio.filePath.Length) break;
                    goto Next;
                }
                GlobalData[i].File.FileName = fio.files[i];
                GlobalData[i].File.FullPath = fio.filePath[i];
                GlobalData[i].File.DateOfCreation = fio.fileCreationTime[i];
            }
            Array.Resize<ust_LogFile>(ref GlobalData, j);
        }
        public void gdDebug_output()
        {
            for (int i = 0; i < GlobalData.Length; i++) Console.WriteLine("#{0} File: {1} Created: {2}", i+1, GlobalData[i].File.FileName, GlobalData[i].File.DateOfCreation);
        }
        #endregion

        #region ExcelIO segment

        public ust_LogSmeta[] getLogFileData(string fullPath)
        {
            if (!File.Exists(fullPath)) return null;
            ExcelIO eio = new ExcelIO(fullPath);
            eio.Open();
            
            ust_LogSmetaRegion[] lsr = getSmetaLogRegion(eio);
            ust_LogSmeta[] lSmeta = new ust_LogSmeta[lsr.Length];
            
            for (int smCnt = 0; smCnt < lSmeta.Length; smCnt++)
            {
                lSmeta[smCnt].Description.Smeta = getUstSmetaData(lsr[smCnt].adrSmeta, eio);
                {
                    string[] eventsAddress = getEventsAdress(lsr[smCnt], eio);
                    lSmeta[smCnt].Data = getEventData(lsr[smCnt], eio);
                    lSmeta[smCnt].Description.Loaded = isSmetaLoaded(lsr[smCnt].adrSmeta, eio);
                }
            }

            eio.Quit();
            return lSmeta;
        }

        public void GetLogFileGlobalData()
        {

            int Count = GlobalData.Length;
            ExcelIO eio = new ExcelIO();
            Excel.Range[] tRanges;
  
            for (int i = 0; i < Count; i++)
            {
                if (!File.Exists(GlobalData[i].File.FullPath)) return;
                eio.Open(GlobalData[i].File.FullPath);
                

                
                eio.CloseWB();
            }
        }

        #endregion

        #region Events data segment

        public string[] getEventsAdresses(ExcelIO eio)
        {
            string eventMark = "Код ГС (SysID)";
            const int smetaColumn = 3;
            //string[] tAddr = eio.search_addr_array(eventMark, column: smetaColumn);
            string[] tAddr = eio.find(eventMark, 1, smetaColumn, eio.maxRows, smetaColumn);

            for (int i = 0; i < tAddr.Length; i++)
            {

                tAddr[i] = eio.getAddress(eio.getRow(tAddr[i]) + 1, 1);
            }

            return tAddr;
        }
        
        public string[] getEventsAdress(ust_LogSmetaRegion lsr, ExcelIO eio)
        {
            string[] AddrPool = eio.search_addr_exception_array("", column: 1, startFrom: lsr.eventRow, Finish: lsr.endRow);
            return AddrPool;
        }

        #region commented code #2
        /*
        public string[] getEventsAdress(string smetaAddress, ExcelIO eio)
        {
            const string eventMark = "Код ГС (SysID)";
            const int eventColumn = 3;

            string tAddr = eio.search_addr(eventMark, column: eventColumn, startFrom: eio.getRow(smetaAddress));
            tAddr = eio.getRelativeAddress(tAddr, 1);
            //tAddr = eio.getAddress(eio.getRow(tAddr), 1);

            string[] AddrPool = eio.search_addr_exception_array("", column: 1, startFrom: eio.getRow(tAddr));
            return AddrPool;
        }

        public string[] getErrorEventsAddr(string smetaAddress, ExcelIO eio)
        {
            string[] events = getEventsAdress(smetaAddress, eio);
            const string errorMark = "Ошибка";
            string[] errorEvents = new string[1];
            int cnt = 0;

            for (int i = 0; i < events.Length; i++)
            {
                if ((eio.getCellValue(eio.getRelativeAddress(events[i], col: 1))).Contains(errorMark))
                {
                    Array.Resize<string>(ref errorEvents, ++cnt);
                    errorEvents[cnt - 1] = events[i];
                }
            }

            return errorEvents;
        }*/
        //не используется
        #endregion

        public ust_LogSmetaData[] getEventData(ust_LogSmetaRegion lsr, ExcelIO eio)
        {
            //int lng = lsr.endRow - lsr.eventRow;
            ust_LogSmetaData[] lsd = new ust_LogSmetaData[1];
            int cnt = 0;

            for (int i = lsr.eventRow; i < lsr.endRow; i++)
            {
                if (!eio.isEmpty(i, 1))
                {
                    Array.Resize<ust_LogSmetaData>(ref lsd, ++cnt);
                    lsd[cnt - 1] = getEventData(i, eio);
                }
            }

            return lsd;
        }

        public ust_LogSmetaData getEventData(int Row, ExcelIO eio)
        {
            if (eio.getCellValue(Row, 1) == "") return new ust_LogSmetaData();
            /*
             * Колонки:
             * 1 - п/п
             * 2 - Событие
             * 3 - Код ГС
             * 4 - Код 1С
             * 5 - Наименование
             * 6 - Описание
             */
            ust_LogSmetaData lsd;

            lsd.ppNumber = eio.getCellValue(Row, 1);
            lsd.Event = eio.getCellValue(Row, 2);
            lsd.SysID = eio.getCellValue(Row, 3);
            lsd.Code1C = eio.getCellValue(Row, 4);
            lsd.Name = eio.getCellValue(Row, 5);
            lsd.Description = eio.getCellValue(Row, 6);

            return lsd;
        }

        #endregion

        #region ExcelIO getData segment

        public string[] getSmetaDataLineAdresses(ExcelIO eio)
        {
            string[] smetaMark = { "Загружен", "Не загружен" };
            const int smetaColumn = 3;
            string[] tAddr = eio.search_addr_array(smetaMark, column: smetaColumn);

            for (int i = 0; i < tAddr.Length; i++)
            {
                tAddr[i] = eio.getAddress(eio.getRow(tAddr[i]), 1);
            }

            return tAddr;
        }

        public bool isSmetaLoaded(string smetaAddress_row, ExcelIO eio)
        {
            const string Loaded = "Загружен";
            //const string notLoaded = "Не загружен";
            const int loadColumn = 3;

            string tAddr = eio.getAddress(eio.getRow(smetaAddress_row), loadColumn);
            tAddr = eio.getCellValue(tAddr);

            if (tAddr == Loaded) return true;
            else return false;
        }

        #endregion

        #region User structures segment

        public ust_LogSmetaRegion[] getSmetaLogRegion(ExcelIO eio)
        {
            return getSmetaLogRegion(getSmetaDataLineAdresses(eio), eio);
        }

        public ust_LogSmetaRegion[] getSmetaLogRegion(string[] smetaAddresses, ExcelIO eio)
        {
            int finalRow = eio.maxRows;
            string titleMark = "Файл";
            int titleColumn = 2;

            ust_LogSmetaRegion[] lsr = new ust_LogSmetaRegion[1];
            int cnt = 0;
            string[] eAdresses = getEventsAdresses(eio);
            string[] titleAdresses = eio.search_addr_array(titleMark, column: titleColumn);
            //удаляем первый адрес не привязанный к сметам
            List<string> lAdresses = new List<string>(eAdresses);
            lAdresses.RemoveAt(0);
            eAdresses = null;
            System.GC.Collect();
            eAdresses = lAdresses.ToArray();

            if(smetaAddresses.Length != eAdresses.Length) return null;

            for (int i = 0; i < smetaAddresses.Length; i++)
            {
                Array.Resize<ust_LogSmetaRegion>(ref lsr, ++cnt);
                lsr[cnt-1].startRow = eio.getRow(smetaAddresses[i]);
                lsr[cnt - 1].eventRow = eio.getRow(eAdresses[i]);
                if (i == (smetaAddresses.Length - 1)) lsr[cnt - 1].endRow = finalRow - 1;
                else lsr[cnt - 1].endRow = (eio.getRow(titleAdresses[i + 1]) - 1);

                lsr[cnt - 1].adrEvent = eAdresses[i];
                lsr[cnt - 1].adrSmeta = smetaAddresses[i];
            }

            return lsr;
        }

        #region commented code #1
        /*
        public ust_LogSmetaData getUstLogSmetaData(string[] strData)
        {
            /*
             * Колонки:
             * 1 - п/п
             * 2 - Событие
             * 3 - Код ГС
             * 4 - Код 1С
             * 5 - Наименование
             * 6 - Описание
             *
            ust_LogSmetaData data;

            data.Code1C = strData[3];
            data.Description = strData[5];
            data.Event = strData[1];
            data.Name = strData[4];
            data.ppNumber = strData[0];
            data.SysID = strData[2];

            return data;
        }

        public ust_LogSmetaData getUstLogSmetaData(string eventAddress, ExcelIO eio)
        {
            /*
             * Колонки:
             * 1 - п/п
             * 2 - Событие
             * 3 - Код ГС
             * 4 - Код 1С
             * 5 - Наименование
             * 6 - Описание
             *
            string[] strData = eio.getRangeData(eventAddress, Start: 1, Finish: 6);
            ust_LogSmetaData data;

            data.Code1C = strData[3];
            data.Description = strData[5];
            data.Event = strData[1];
            data.Name = strData[4];
            data.ppNumber = strData[0];
            data.SysID = strData[2];

            return data;
        }

        public ust_LogSmetaData[] getUstLogSmetaDataArray(string[] eventsAddress, ExcelIO eio)
        {
            ust_LogSmetaData[] data = new ust_LogSmetaData[1];
            string[] strData;
            int cnt = 0;

            if (eventsAddress.Length == 1 && eventsAddress[0] == null) return null;

            for (int i = 0; i < eventsAddress.Length; i++)
            {
                strData = null;
                if (eventsAddress[i] != null)
                {
                    strData = eio.getRangeData(eventsAddress[i]);
                    Array.Resize<ust_LogSmetaData>(ref data, cnt + 1);

                    data[cnt].Code1C = strData[3];
                    data[cnt].Description = strData[5];
                    data[cnt].Event = strData[1];
                    data[cnt].Name = strData[4];
                    data[cnt].ppNumber = strData[0];
                    data[cnt].SysID = strData[2];
                    cnt++;
                }
            }
            System.GC.Collect();

            return data;
        }

        public ust_Smeta getUstSmetaData(string[] Row)
        {
            /*
             * Колонки:
             * 2 - Имя файла
             * 3 - Статус         
             * 4 - Код сметы
             * 5 - Проект
             * 6 - Объект строительства
             * 8 - Наименование сметы
             * 10 - Номер
             * 15 - Дата загрузки - не реаизовано
             *
            //const string Status = "Загружен";
            ust_Smeta Smeta = new ust_Smeta();

            //if (Row.Length != 15 || Row[2] != Status) return Smeta;

            Smeta.FileName = Row[1];
            Smeta.Status = Row[2];
            Smeta.Code = Row[3];
            Smeta.Project = Row[4];
            Smeta.Object = Row[5];
            Smeta.Name = Row[7];
            Smeta.Number = Row[9];

            return Smeta;
        }*/ //Не используется
        #endregion

        public ust_Smeta getUstSmetaData(string smetaAddress, ExcelIO eio)
        {
            string[] Row = eio.getRangeData(smetaAddress, false, Start: 1, Finish: 15);
            ust_Smeta Smeta = new ust_Smeta();

            Smeta.FileName = Row[1];
            Smeta.Status = Row[2];
            Smeta.Code = Row[3];
            Smeta.Project = Row[4];
            Smeta.Object = Row[5];
            Smeta.Name = Row[7];
            Smeta.Number = Row[9];

            return Smeta;
        }
        #endregion
    }
}

#region Temporary
/*
       public void GetLogFileGlobalData()
        {

            const string smetaMark = "Файл";
            const string eventMark = "Код ГС (SysID)";
            const int smetaCol = 2;
            const int eventCol = 3;

            int Count = GlobalData.Length;
            ExcelIO eio = new ExcelIO();
            Excel.Range[] tRanges;
            string[] tAddresses;
            string[] tmpArray;
            string[] tmpEventsRows;
            string tSmetaAddr = "";
            string tEventsAddr = "";

            for (int i = 0; i < Count; i++)
            {
                if (!File.Exists(GlobalData[i].File.FullPath)) return;
                eio.Open(GlobalData[i].File.FullPath);
                tAddresses = null;
                tRanges = null;
                tEventsAddr = null;
                tmpArray = null;
                System.GC.Collect();


                tSmetaAddr = eio.search_addr(smetaMark, -1, 2, true, (tSmetaAddr == "") ? 1 : eio.getRow(eio.getRelativeAddress(tSmetaAddr, 1)));
                
                if (tSmetaAddr == "") break;

                tmpArray = eio.getRangeData(eio.getRelativeAddress(tSmetaAddr, 1),true, 1);
                GlobalData[i].Body.Description.Smeta = getUstSmetaData(tmpArray);

                tEventsAddr = eio.getRelativeAddress(eio.search_addr(eventMark, -1, 3, true, eio.getRow(eio.getRelativeAddress(tSmetaAddr, 1))), 1);

                tAddresses = eio.search_addr_exception_array("", 0, 1, true, eio.getRow(tSmetaAddr));
                GlobalData[i].Body.Data = new ust_LogSmetaData[tAddresses.Length];
                
                for(int j = 0; j < tAddresses.Length;j++)
                {
                    tmpEventsRows = eio.getRangeData(tAddresses[j], true, 1, 6);
                    
                    GlobalData[i].Body.Data[j] = getUstLogSmetaData(tmpEventsRows);


                    tmpEventsRows = null;
                    System.GC.Collect();
                }

                eio.CloseWB();
            }
        }
*/
#endregion