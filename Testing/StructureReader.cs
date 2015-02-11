using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    /* ust_LogFile
     *  File (ust_LogFileDescription)
     *      FullPath (string)
     *      FileName (string)
     *      DateOfCreation (DateTime)
     *  
     *  Body (ust_LogSmeta)
     *      Description (ust_LogSmetaDescription)
     *          Smeta (ust_Smeta)
     *          LoadTime (DateTime)
     *          Loaded (bool)
     *      Data (ust_LogSmetaData[])       
     * 
     */

    class structureReader
    {
        #region Variables

        private ust_lstruct[] data;
        public ust_LogSmeta[] smetalog { get; private set; }
        private bool isRead;

        private ExcelIO eio;

        #endregion

        #region Constructors
        //public void structureReader() { }
        public structureReader(ust_lstruct[] structureBuilderData, ExcelIO Excel)
        {
            data = structureBuilderData;
            eio = Excel;
        }

        #endregion

        #region User public methods

        public void Connect(ust_lstruct[] structureBuilderData, ExcelIO Excel)
        {
            clear();
            data = structureBuilderData;
            eio = Excel;
        }

        public void Read()
        {
            smetalog = new ust_LogSmeta[data.Length];

            for (int i = 0; i < data.Length; i++)
            {
                smetalog[i].Description.Smeta = getSmetaLine(data[i].smeta);
                setSmetaDescriptionArguments(ref smetalog[i].Description);
                smetalog[i].Data = getEventsArray(i);
            }
            isRead = true;
        }

        /*
        public ust_LogSmeta[] getLogSmeta()
        { return smetalog; }*/

        #endregion

        #region Service

        private void clear()
        {
            data = null;
            smetalog = null;
            isRead = false;
            eio = null;

            System.GC.Collect();
        }

        #endregion

        #region User structures writing methods

        private void setSmetaDescriptionArguments(ref ust_LogSmetaDescription lsd)
        {
            lsd = isLoaded(lsd);
            if (lsd.Loaded) lsd.LoadTime = DateTime.Parse(lsd.Smeta.Time);
        }

        #endregion

        #region Read from Excel methods

        private ust_LogSmetaData[] getEventsArray(int index)
        {
            if (data[index].events == null) return null;

            int lng = data[index].events.Length;
            int Row = eio.getRow(data[index].events[0]);
            ust_LogSmetaData[] evs = new ust_LogSmetaData[lng];

            for (int i = 0; i < lng; i++) evs[i] = getEventLine(Row++);

            return evs;
        }

        private ust_LogSmetaData getEventLine(string adr)
        {
            int Row = eio.getRow(adr);

            return getEventLine(Row);
        }

        private ust_LogSmetaData getEventLine(int Row)
        {
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

        private ust_Smeta getSmetaLine(string adr)
        {
            ust_Smeta Smeta = new ust_Smeta();
            int Row = eio.getRow(adr);

            //15 - колонка дата загрузки

            Smeta.FileName = eio.getCellValue(Row, 2);
            Smeta.Status = eio.getCellValue(Row, 3);
            Smeta.Code = eio.getCellValue(Row, 4);
            Smeta.Project = eio.getCellValue(Row, 5);
            Smeta.Object = eio.getCellValue(Row, 6);
            Smeta.Name = eio.getCellValue(Row, 8);
            Smeta.Number = eio.getCellValue(Row, 10);
            Smeta.Time = eio.getCellValue(Row, 15);

            return Smeta;
        }

        #endregion

        #region Check methods

        private ust_LogSmetaDescription isLoaded(ust_LogSmetaDescription lsd)
        {
            string loaded = "Загружен";
            if (lsd.Smeta.Status == loaded) lsd.Loaded = true;
            else lsd.Loaded = false;
            return lsd;
        }

        #endregion

    }
}
