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
        static bool isDigit(char c)
        {
            int r;
            return Int32.TryParse(Convert.ToString(c), out r);
        }

        static int getLetterNumber(char C)
        {
            char[] L = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
            List<char> letters = new List<char>(L);
            C = Char.ToUpper(C);
            return (letters.IndexOf(C) + 1);
        }

        static int getColumn(string addr)
        {
            string column = getColumnString(addr);
            int lng = column.Length;

            if (lng == 0 || lng > 3) return 0;

            switch (lng)
            {
                case 1:
                    return 

            }


        }

        static string getColumnString(string addr)
        {
            char[] src = addr.ToCharArray();
            StringBuilder result = new StringBuilder();

            for (int i = 0; i < addr.Length; i++)
            {
                if (!isDigit(src[i])) result.Append(src[i]);
                else break;
            }

            return result.ToString();
        }

        static int getRow(string addr)
        {
            char[] src = addr.ToCharArray();
            StringBuilder result = new StringBuilder();

            for (int i = addr.Length - 1; i >= 0; i--)
            {
                if (isDigit(src[i]))
                {
                    if (result.Length == 0) result.Append(src[i]);
                    else result.Insert(0, src[i]);
                }
            }
            
            return Convert.ToInt32(result.ToString());
        }

        static string toAddress(int Row, int Column) //Доделать
        {
            //AAB = 704
            //26
            // 1 - 26
            string[] L = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
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

            /*
            ExcelIO eio = new ExcelIO(@"C:\WORK\Log_2015-01-16_15-06-27.xlsx");
            eio.Open();
            string[] addr = eio.find("", "A1", "A245");
            eio.Quit();
            /*
            ErrorLog el = new ErrorLog();
            ust_LogSmeta[] ustLS = el.getLogFileData(@"C:\WORK\Log_2015-01-16_15-06-27.xlsx");*/
            int x = 704;
            int y = 26;
            int c = x / y;
            int d = x % y;
            string adr = toAddress(12345, 729);
            int row = getRow(adr);
            string col = getColumnString(adr);

            Console.WriteLine("Address: {0}\nRow: {1}\nColumn: {2}", adr, row, col);

            Console.ReadLine();
        }
    }
}
