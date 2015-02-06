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
        
#region Constants
        
        private const string eventTag = "Код ГС (SysID)";
        private const string smetaTag = "Смета";
        private const int smetaCol = 9;
        private const int eventsCol = 4;

#endregion
                
#endregion

#region Constructors

        public LogStructure() 
        {
            data = new ust_lstruct[1];
        }

#endregion

#region Public Methods

        public void addSmeta(string addr)
        {
            Array.Resize<ust_lstruct>(ref data, ++count);
            current = count - 1;
            
            data[current].smeta = addr;
        }

        public void addEvent(string addr)
        {
            if (data[current].events == null) data[current].events = new string[1];
            else Array.Resize(ref data[current].events, data[current].events.Length + 1);

            data[current].events[data[current].events.Length - 1] = addr;
        }

        private string adrSmetaStartFnd(ExcelIO eio)
        {
            string tag = "Файл";
            string adrStart = "B1";
            string adrFinish = "B100";

            string adr = eio.find_once(tag, adrStart, adrFinish);
            return eio.getRelativeAddress(adr, col: 7);
        }

        public void buildStructure(ExcelIO eio)
        {
            
        }

#endregion


    }
}
