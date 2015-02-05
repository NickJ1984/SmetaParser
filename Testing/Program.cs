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

        static bool isLetter(char c)
        {
            char r;
            return char.TryParse(Convert.ToString(c), out r);
        }

        static char getLetterCharacter(int number)
        {
            if (number < 1 || number > 26) return '-';
            char[] L = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
            return L[number - 1];
        }

        static int getLetterNumber(char C)
        {
            char[] L = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
            List<char> letters = new List<char>(L);
            C = Char.ToUpper(C);
            int index = letters.IndexOf(C);
            if (index >= 0) return index + 1;
            else return index;
        }

        static int getColumn(string addr)
        {
            char[] column = (getColumnString(addr)).ToCharArray();
            int lng = column.Length;

            if (lng == 0 || lng > 3) return 0;

            int units;
            int ten;
            int hundreds;

            switch (lng)
            {
                case 1:
                    return getLetterNumber(column[0]);

                case 2:
                    units = getLetterNumber(column[1]);
                    ten = getLetterNumber(column[0]) * 26;
                    return ten + units;

                case 3:
                    units = getLetterNumber(column[2]);
                    ten = getLetterNumber(column[1]) * 26;
                    hundreds = getLetterNumber(column[0]) * 26 * 26;
                    return ten + units + hundreds;
                
                default:
                    return 0;
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
            string adr = toAddress(12345, 2731);
            int row = getRow(adr);
            string col = getColumnString(adr);
            int colNum = getColumn(adr);

            Console.WriteLine("Address: {0}\nRow: {1}\nColumn: {2}\nColumn number: {3}", adr, row, col, colNum);
            Console.WriteLine("Letter: C\nLetter: {0}\nLetter number: {1}", getLetterCharacter(30),getLetterNumber('!'));

            Console.ReadLine();
        }
    }
}
