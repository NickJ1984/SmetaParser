using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    class structureReader
    {
        #region Variables

        private ust_lstruct[] data;
        private ust_LogFile[] log { get; private set; }
        private ExcelIO eio;

        #endregion

        #region Constructors
        //public void structureReader() { }
        public void structureReader(ust_lstruct[] structureBuilderData, ExcelIO Excel)
        {
            data = structureBuilderData;
            eio = Excel;
        }

        #endregion


        private void process()
        {

        }

        #region User structures writing methods

        private ust_LogSmetaData[] getEventsArray(int index)
        {
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

        private void isLoaded(ust_LogSmetaDescription lsd)
        {

        }

        #endregion

        #region Read from Excel methods

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

            Smeta.FileName = eio.getCellValue(Row, 2);
            Smeta.Status = eio.getCellValue(Row, 3);
            Smeta.Code = eio.getCellValue(Row, 4);
            Smeta.Project = eio.getCellValue(Row, 5);
            Smeta.Object = eio.getCellValue(Row, 6);
            Smeta.Name = eio.getCellValue(Row, 8);
            Smeta.Number = eio.getCellValue(Row, 10);

            return Smeta;
        }

        #endregion

    }
}
