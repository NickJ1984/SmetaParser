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
        /*
        public ust_Smeta getLSmetaData(string[] Row)
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
             * 15 - Дата загрузки
         
            const string Status = "Загружен";

            return null;

        }*/
        

        static string toAddress(int Row, int Column) //Доделать
        {
            //AAB = 704
            //26
            // 1 - 26
            string[] Letters = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            const int level = 26;

            int div = Column / level;
            int units = (Column > 26) ? Column % level : Column;
            int ten = (div > 26) ? div % level : div;
            int hundreds = (div > 26) ? div / 26 : 0;

            string result =
                ((hundreds == 0) ? "" : Letters[hundreds - 1]) +
                ((ten == 0) ? "" : Letters[ten - 1]) +
                Letters[units-1] +
                Convert.ToString(Row);

            return result;
        }

        static void Main(string[] args)
        {

            ExcelIO eio = new ExcelIO(@"C:\WORK\Log_2015-01-16_15-06-27.xlsx");
            eio.Open();
            //string[] addr = eio.find("", "A1", "A245");
            string[] addr = eio.find_exception("", "A25", "A136");
            sup.writelnArray(addr);

            eio.Quit();
            /*
            ErrorLog el = new ErrorLog();
            ust_LogSmeta[] ustLS = el.getLogFileData(@"C:\WORK\Log_2015-01-16_15-06-27.xlsx");*/
            
            Console.ReadLine();
        }
    }
}
