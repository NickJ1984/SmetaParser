using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApplication1
{
    class cellAddress
    {
        #region Variables

        public int Row {get; private set;}
        public int ColumnI { get; private set; }
        public string ColumnS { get; private set; }
        public string Address { get; private set; }
        private object value;
        private bool isEmpty;

        #endregion

        #region Constructors

        public cellAddress() { }

        public cellAddress(int srcRow, int srcColumn)
        {
            if (rowValidCheck(srcRow) && columnValidCheck(srcColumn))
            {
                addressSet(srcRow, srcColumn);
            }
        }

        public cellAddress(int srcRow, string srcColumn)
        {
            if (rowValidCheck(srcRow) && columnValidCheck(srcColumn))
            {
                addressSet(srcRow, srcColumn);
            }
        }

        public cellAddress(string srcAddress)
        {
            addressSetValue(srcAddress);
        }

        public cellAddress(cellAddress cellAdr)
        {
            addressSet(cellAdr.Row, cellAdr.ColumnS);
        }

        #endregion

        #region Class interaction methods

        #region Methods overrides

        public override string ToString()
        {
            return Address;
        }

        #endregion

        #region Operators overrides

        #endregion

        public void Copy(cellAddress ca)
        {

        }

        public void Clear()
        {

        }

        #endregion

        #region Address convertation methods

        #region Column

        private string getColumnLetters(string addr)
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

        private int getColumn(string Column)
        {
            char[] column = (getColumnLetters(Column)).ToCharArray();
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

        private string getColumn(int Column)
        {
            //AAB = 704
            //26
            // 1 - 26
            if (Column < 1) return null;

            string[] Letters = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            const int level = 26;

            int div = Column / level;
            int units = (Column > 26) ? Column % level : Column;
            int ten = (div > 26) ? div % level : div;
            int hundreds = (div > 26) ? div / 26 : 0;

            string result =
                ((hundreds == 0) ? "" : Letters[hundreds - 1]) +
                ((ten == 0) ? "" : Letters[ten - 1]) +
                Letters[units - 1];

            return result;
        }

        #endregion

        #region Row

        private int getRow(string addr)
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
                else break;
            }

            return Convert.ToInt32(result.ToString());
        }

        #endregion

        #region Address

        private string getAddressInitial(string addr)
        {
            if (addr == null) return null;

            return "A" + getRow(addr);
        }

        private string getAddressInitial(int Row)
        {
            if (Row <= 0) return null;
            return "A" + Row;
        }

        private string getAddress(int Row, int Column)
        {
            //AAB = 704
            //26
            // 1 - 26
            if (Row < 1 || Column < 1) return null;

            string[] Letters = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            const int level = 26;

            int div = Column / level;
            int units = (Column > 26) ? Column % level : Column;
            int ten = (div > 26) ? div % level : div;
            int hundreds = (div > 26) ? div / 26 : 0;

            string result =
                ((hundreds == 0) ? "" : Letters[hundreds - 1]) +
                ((ten == 0) ? "" : Letters[ten - 1]) +
                Letters[units - 1] +
                Convert.ToString(Row);

            return result;
        }

        private string getAddress(int Row, string Column)
        {
            return getAddress(Row, getColumn(Column));
        }

        public string getAddress(Excel.Range rng)
        { return rng.Address; }

        private string getRelativeAddress(string addr, int row = 0, int col = 0)
        {
            int cRow = getRow(addr);
            int cColumn = getColumn(addr);

            if (cRow + row < 1) return null;
            if (cColumn + col < 1) return null;

            return getAddress(cRow + row, cColumn + col);
        }

        #endregion

        #endregion
        
        #region Class mechanics

        #region column methods

        private void columnSet(int index)
        {
            ColumnI = index;
            ColumnS = getColumn(index);
        }

        private void columnSet(string col)
        {
            ColumnI = getColumn(col);
            ColumnS = col;
        }

        #endregion

        #region row methods

        private void rowSet(int index)
        {
            Row = index;
        }

        #endregion

        #region address methods

        private void addressSet(int Row, string Column)
        {
            columnSet(Column);
            rowSet(Row);
            addressChange();
        }

        private void addressSet(int Row, int Column)
        {
            columnSet(Column);
            rowSet(Row);
            addressChange();
        }

        private void addressSet(string adr)
        {
            string col = getColumnLetters(adr);
            int r = getRow(adr);
            columnSet(col);
            rowSet(r);
            addressChange();
        }

        private void addressChange()
        {
            //Address = getAddress(Row, ColumnS);
            Address = ColumnS + Convert.ToString(Row);
        }

        #endregion

        #endregion

        #region public user methods

        #region Set

        public void columnSetValue(int index)
        {
            if (index > 0 && columnChangeCheck(index))
            {
                columnSet(index);
                addressChange();
            }
        }

        public void columnSetValue(string index)
        {
            if (index != "" && columnChangeCheck(index))
            {
                columnSet(index);
                addressChange();
            }
        }

        public void rowSetValue(int index)
        {
            if (index > 0 && rowChangeCheck(index))
            {
                rowSet(index);
                addressChange();
            }
        }

        public void addressSetValue(string adr)
        {
            if (adr == "") return;
            adr = addressDollarClear(adr);
            
            if(addressChangeCheck(adr)) addressSet(adr);
        }

        #endregion

        #endregion

        #region Service methods

        private string addressDollarClear(string addr)
        {
            return addr.Replace("$", "");
        }

        private bool isDigit(char c)
        {
            int r;
            return Int32.TryParse(Convert.ToString(c), out r);
        }

        private bool isLetter(char c)
        {
            char r;
            return char.TryParse(Convert.ToString(c), out r);
        }

        private char getLetterCharacter(int number)
        {
            if (number < 1 || number > 26) return '-';
            char[] L = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
            return L[number - 1];
        }

        private int getLetterNumber(char C)
        {
            char[] L = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
            List<char> letters = new List<char>(L);
            C = Char.ToUpper(C);
            int index = letters.IndexOf(C);
            if (index >= 0) return index + 1;
            else return index;
        }

        #endregion

        #region Check methods

        private bool checkStartFinishRegion(int Start, int Finish)
        {
            return
                (Start <= Finish
                && Start > 0
                && Finish > 0
                && Start < Finish) ? true : false;
        }

        private bool columnChangeCheck(string col)
        {
            if (ColumnS != col) return true;
            else return false;
        }

        private bool columnChangeCheck(int col)
        {
            if (ColumnI != col) return true;
            else return false;
        }

        private bool columnValidCheck(int col)
        {
            if (col > 0) return true;
            else return false;
        }

        private bool columnValidCheck(string col)
        {
            if (col.Length > 0 && col.Length <= 3) return true;
            else return false;
        }

        private bool rowChangeCheck(int r)
        {
            if (Row != r) return true;
            else return false;
        }

        private bool rowValidCheck(int r)
        {
            if (r > 0) return true;
            else return false;
        }

        private bool addressChangeCheck(string adr)
        {
            if (Address != adr) return true;
            else return false;
        }

        #endregion

    }
}