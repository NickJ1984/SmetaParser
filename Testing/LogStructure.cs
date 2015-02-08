using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    enum uen_flagsStructure { Sm = 1, Ev}

    struct ust_lstruct
    {
        public string smeta;
        public string[] events;
    }

    class LogStructure
    {
#region Variables

        private ust_lstruct[] data;
        private int count;
        private int current;
        private string adrSmEnd;
        private string adrEvEnd;
        private ExcelIO eio;

        #region Constants
        
        private const string eventTag = "Код ГС (SysID)";
        //private const string smetaTag = "Смета";
        private const string smetaTag = ".xml";
        //private const int smetaCol = 9;
        private const int smetaCol = 2;
        private const int eventsCol = 3;

        #endregion
                
#endregion

#region Constructors

        public LogStructure(ExcelIO source) 
        {
            data = new ust_lstruct[1];
            eio = source;
        }

#endregion

#region Public Methods

        #region Add methods

        private void addSmeta(string addr)
        {
            Array.Resize<ust_lstruct>(ref data, ++count);
            current = count - 1;
            
            data[current].smeta = addr;
        }

        private void addEvent(string addr)
        {
            if (data[current].events == null) data[current].events = new string[1];
            else Array.Resize(ref data[current].events, data[current].events.Length + 1);

            data[current].events[data[current].events.Length - 1] = addr;
        }

        private void addEvent(string[] addr)
        {
            if (data[current].events == null)
            {
                data[current].events = addr;
                return;
            }
            else Array.Resize(ref data[current].events, data[current].events.Length + addr.Length);

            for (int i = 0; i < addr.Length; i++) data[current].events[data[current].events.Length + i] = addr[i];
        }

        #endregion

        #region Adress methods

        private string adrStartFnd()
        {
            string tag = "Файл";
            string adrStart = "B1";
            string adrFinish = "B100";

            string adr = eio.find_once(tag, adrStart, adrFinish);
            return adr;
            //return eio.getRelativeAddress(adr, col: 7);
        }

        private string getAdrSm(string adr)
        {
            //return eio.find_once(smetaTag, eio.getRelativeAddress(adr, 1), adrSmEnd);
            int row = eio.getRow(adr) + 1;
            string address = eio.getAddress(row, smetaCol);
            return eio.find_once(smetaTag, address, adrSmEnd);
        }

        private string getAdrEv(string adrSm)
        {
            string adr = eio.find_once(eventTag, eio.getAddress(eio.getRow(adrSm), eventsCol), adrEvEnd);
            adr = eio.getAddress(eio.getRow(adr) + 1, "A");
            
            return adr;
        }

        #endregion

        public void buildStructure()
        {
            if (!eio.isOpen)
            {
                Console.WriteLine("Excel не открыт");
                return;
            }

            adrEvEnd = eio.getAddress(eio.maxRows, eventsCol);
            adrSmEnd = eio.getAddress(eio.maxRows, smetaCol);
            
            string adrSmStart = adrStartFnd();
            string adrSm = getAdrSm(eio.getRelativeAddress(adrSmStart,-1));
            
            string adrEv = null;

            while (adrSm != null)
            {
                addSmeta(eio.getAddressInitial(adrSm));
                adrEv = getAdrEv(adrSm);
                addEvent(eio.find_exception("", adrEv, "A" + eio.maxRows));

                adrSm = getAdrSm(adrSm);
            }

        }

#endregion


    }
}
