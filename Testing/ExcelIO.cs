using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApplication1
{
    #region User structures

    struct ust_Cell
    {
        int Row;
        int Column;
    }

    #endregion

    class ExcelIO
    {
        #region Variables
        public string path { get; private set; }

        public int maxColumns { get; private set; }
        public int maxRows { get; private set; }
        private bool isOpen, isAppExcelOpen;
        private Excel.Application appExcel;
        private Excel.Workbook wbExcel;
        private Excel.Worksheet wsExcel;
        #endregion

        #region Constructors
        public ExcelIO()
        {
            maxColumns = 0;
            maxRows = 0;
            isOpen = false;
            isAppExcelOpen = false;
        }
        public ExcelIO(string filePath)
            : base()
        {
            path = filePath;
        }
        #endregion

        #region File operations

        public void Open(string fullpath = "")
        {
            if (isOpen) return;

            if (fullpath == "") { if (path == "") return; }
            else path = fullpath;
            if (!File.Exists(path)) return;
            if (!isAppExcelOpen)
            {
                appExcel = new Excel.Application();
                isAppExcelOpen = true;
            }
            wbExcel = appExcel.Workbooks.Open(path, ReadOnly: true);
            wsExcel = (Excel.Worksheet)wbExcel.Sheets[1];
            maxRows = wsExcel.UsedRange.Rows.Count;
            maxColumns = wsExcel.UsedRange.Columns.Count;
            isOpen = true;
        }

        public void Quit()
        {
            object missingObj = System.Reflection.Missing.Value;

            wbExcel.Close(false, missingObj, missingObj);
            appExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(appExcel);
            appExcel = null;
            wbExcel = null;
            wsExcel = null;
            System.GC.Collect();
            isOpen = false;
            isAppExcelOpen = false;
        }

        public void CloseWB()
        {
            object missingObj = System.Reflection.Missing.Value;
            wbExcel.Close(false, missingObj, missingObj);
            wbExcel = null;
            wsExcel = null;
            System.GC.Collect();
            isOpen = false;
        }

        #endregion

        #region Address convertation methods

        #region Column

        public string getColumnLetters(string addr)
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

        public int getColumn(string addr)
        {
            char[] column = (getColumnLetters(addr)).ToCharArray();
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

        #endregion

        #region Row

        public int getRow(string addr)
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

        public string getAddressInitial(string addr)
        {
            if (addr == null) return null;

            return "A" + getRow(addr);
        }

        public string getAddressInitial(int Row)
        {
            if (Row <= 0) return null;
            return "A" + Row;
        }

        public string getAddress(int Row, int Column)
        {
            //AAB = 704
            //26
            // 1 - 26
            if (Row < 1 || Column < 1) return null;

            string[] Letters = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};
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

        public string getAddress(Excel.Range rng)
        { return rng.Address; }

        public string getRelativeAddress(string addr, int row = 0, int col = 0)
        {
            int cRow = getRow(addr);
            int cColumn = getColumn(addr);

            if ((cRow + row) > maxRows || (cRow + row) < 1) return null;
            if ((cColumn + col) > maxColumns || (cColumn + col) < 1) return null;

            return getAddress(cRow + row, cColumn + col);
        }

        #endregion

        public int[] getRowColumn(string addr)
        { return new int[2] { (wsExcel.Range[addr] as Excel.Range).Row, (wsExcel.Range[addr] as Excel.Range).Column }; }

        #endregion

        #region Check methods

        private bool checkStartFinishRegion(int Start, int Finish)
        {
            return
                (Start <= Finish
                && Start > 0
                && Finish > 0
                && Start <= maxRows
                && Finish <= maxRows) ? true : false;
        }

        public bool isEmpty(string Address)
        {
            if (getCellValue(Address) == "") return true;
            else return false;
        }

        public bool isEmpty(int row, int column) { return isEmpty(getAddress(row, column)); }

        #endregion

        #region Search methods
        
        #region Find method

        public string[] find(string text, string adr1, string adr2, bool ADR = true)
        {
            Excel.Range area = appExcel.get_Range(adr1, adr2);
            Excel.Range firstFind = null;
            Excel.Range currentFind = null;
            int cnt = 0;
            string[] array = new string[1];

            currentFind = area.Find(text);

            while (currentFind != null)
            {
                if (firstFind == null) firstFind = currentFind;
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1) == firstFind.get_Address(Excel.XlReferenceStyle.xlA1)) break;
                
                Array.Resize<string>(ref array, ++cnt);
                if (ADR) array[cnt - 1] = addressDollarClear(currentFind.get_Address(Excel.XlReferenceStyle.xlA1));
                else array[cnt - 1] = currentFind.Text;

                currentFind = area.FindNext(currentFind);
             }

            return array;
        }

        public string[] find(string text, int Row_1, int Column_1, int Row_2, int Column_2, bool ADR = true)
        {
            return find(text, getAddress(Row_1, Column_1), getAddress(Row_2, Column_2), ADR);
        }

        public string[] find(string text, Excel.Range adr1, Excel.Range adr2, bool ADR = true)
        {
            return find(text, adr1.get_Address(Excel.XlReferenceStyle.xlA1), adr2.get_Address(Excel.XlReferenceStyle.xlA1), ADR);
        }

        public string find_once(string Text, string Adr1, string Adr2, Excel.Range StartFrom = null)
        {
            Excel.Range area = appExcel.get_Range(Adr1, Adr2);
            Excel.Range firstFind = null;

            if (StartFrom != null) firstFind = area.Find(Text, StartFrom);
            else firstFind = area.Find(Text);

            return addressDollarClear(firstFind.get_Address());
        }

        public string[] find_exception(string Text, string StartAddress, string FinishAddress)
        {
            int rsaColumn = getColumn(StartAddress);
            int rfaColumn = getColumn(FinishAddress);
            int rsaRow = getRow(StartAddress);
            int rfaRow = getRow(FinishAddress);
            bool Column = true;

            int diffColumns = rfaColumn - rsaColumn;
            int diffRows = rfaRow - rsaRow;
            if (diffRows != 0 && diffColumns != 0) return null;
            else
            {
                if (diffRows != 0) Column = true;
                else if (diffColumns != 0) Column = false;
                else return null;
            }
            
            string adrFinish = find_once(Text, StartAddress, FinishAddress);
            int start = (Column) ? getRow(StartAddress) : getColumn(StartAddress);
            int finish = (Column) ? getRow(adrFinish) - 1 : getColumn(adrFinish);

            string[] adrArray = new string[finish - start + 1];
            int cnt = 0;

            for (int i = start; i <= finish; i++)
            {
                adrArray[cnt++] = getAddress((Column) ? i : rsaRow,
                                             (Column) ? rsaColumn : i);
            }

            return adrArray;
        }

        #endregion

        public int search(string txt, int row = -1, int column = -1, bool precision = true, int startFrom = 1)
        {
            int max = 0;
            string tmp;
            if (row == -1 && column == -1) return -1;
            max = (column == -1) ? maxColumns : maxRows;
            if (startFrom > max) return -1;

            for (int step = startFrom; step < max; step++)
            {
                tmp = getCellValue((row == -1) ? step : row, (column == -1) ? step : column);
                if (precision) { if (tmp == txt) return step; }
                else
                {
                    if (tmp.Length > txt.Length)
                        if (tmp.Contains(txt)) return step;
                }
            }

            return -1;
        }

        public string search_addr(string txt, int row = -1, int column = -1, bool precision = true, bool breakable = false, int startFrom = 1, int finish = 0)
        {
            int max = 0;
            string tmp;
            if (row == -1 && column == -1) return "";
            max = (column == -1) ? maxColumns : maxRows;
            if (startFrom > max) return "";
            if (finish > 0)
            {
                if (checkStartFinishRegion(startFrom, finish))
                {
                    max = finish;
                }
                else return null;
            }

            for (int step = startFrom; step < max; step++)
            {
                tmp = getCellValue((row == -1) ? step : row, (column == -1) ? step : column);
                if (precision) 
                {
                    if (tmp == txt) return (wsExcel.Cells[(row == -1) ? step : row, (column == -1) ? step : column] as Excel.Range).Address;
                    else if (breakable) break;
                }
                else
                {
                    if (tmp.Length > txt.Length)
                    {
                        if (tmp.Contains(txt)) return (wsExcel.Cells[(row == -1) ? step : row, (column == -1) ? step : column] as Excel.Range).Address;
                        else if (breakable) break;
                    }
                    else if (breakable) break;

                }
            }

            return "";
        }

        public string[] search_addr_exception_array(string txt, int row = 0, int column = 0, bool precision = true, bool breakable = true, int startFrom = 1, int Finish = 0)
        {
            int max = 0;
            string tmp;
            int cnt = 0;
            string[] ans = new string[1];

            if (row == 0 && column == 0) return null;
            max = (column == 0) ? maxColumns : maxRows;
            if (startFrom > max) return null;
            if (Finish > 0 && Finish < max && Finish > startFrom) max = Finish;

            for (int step = startFrom; step < max; step++)
            {
                tmp = getCellValue((row == 0) ? step : row, (column == 0) ? step : column);

                if ((tmp != txt && precision) || (tmp.Length >= txt.Length && !tmp.Contains(txt) && !precision))
                {
                    Array.Resize<string>(ref ans, ++cnt);

                    ans[cnt - 1] = (wsExcel.Cells[(row == 0) ? step : row, (column == 0) ? step : column] as Excel.Range).Address;
                }
                else if(breakable) break;
            }
            return ans;
        }


        public string[] search_addr_array(string txt, int row = 0, int column = 0, bool precision = true, int startFrom = 1, int finish = 0)
        {
            int max = 0;
            string tmp;
            int cnt = 0;
            string[] ans = new string[1];

            if (row == 0 && column == 0) return null;
            max = (column == 0) ? maxColumns : maxRows;
            if (startFrom > max) return null;
            if (finish > 0)
            {
                if (checkStartFinishRegion(startFrom, finish))
                {
                    max = finish;
                }
                else return null;
            }
            /*if (row > 0) column++;
            else row++;*/
                
            for (int step = startFrom; step < max; step++)
            {
                tmp = getCellValue((row == 0) ? step : row, (column == 0) ? step : column);
                if (precision)
                {
                    if (tmp == txt)
                    {
                        Array.Resize<string>(ref ans, ++cnt);
                        ans[cnt - 1] = (wsExcel.Cells[(row == 0) ? step : row, (column == 0) ? step : column] as Excel.Range).Address;
                    }
                }
                else
                {
                    if (tmp.Length >= txt.Length)
                    {
                        if (tmp.Contains(txt))
                        {
                            Array.Resize<string>(ref ans, ++cnt);
                            ans[cnt - 1] = (wsExcel.Cells[(row == 0) ? step : row, (column == 0) ? step : column] as Excel.Range).Address;
                        }
                    }
                }
            }
            return ans;
        }

        public string[] search_addr_array(string[] txt, int row = 0, int column = 0, bool precision = true, int startFrom = 1, int finish = 0)
        {
            int max = 0;
            string tmp;
            int cnt = 0;
            string[] ans = new string[1];

            if (row == 0 && column == 0) return null;
            max = (column == 0) ? maxColumns : maxRows;
            if (startFrom > max) return null;
            if (finish > 0)
            {
                if (checkStartFinishRegion(startFrom, finish))
                {
                    max = finish;
                }
                else return null;
            }
            /*if (row > 0) column++;
            else row++;*/

            for (int step = startFrom; step < max; step++)
            {
                tmp = getCellValue((row == 0) ? step : row, (column == 0) ? step : column);
                if (precision)
                {
                    for (int i = 0; i < txt.Length; i++)
                    {
                        if (tmp == txt[i])
                        {
                            Array.Resize<string>(ref ans, ++cnt);
                            ans[cnt - 1] = (wsExcel.Cells[(row == 0) ? step : row, (column == 0) ? step : column] as Excel.Range).Address;
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < txt.Length; i++)
                    {
                        if (tmp.Length >= txt[i].Length)
                        {
                            if (tmp.Contains(txt[i]))
                            {
                                Array.Resize<string>(ref ans, ++cnt);
                                ans[cnt - 1] = (wsExcel.Cells[(row == 0) ? step : row, (column == 0) ? step : column] as Excel.Range).Address;
                            }
                        }
                    }
                }
            }
            return ans;
        }

        #endregion

        #region Range methods

        public Excel.Range[] getRanges(string[] addr, bool fullRange = true)
        {
            if (addr == null) return null;
            
            Excel.Range[] rngResult = new Excel.Range[addr.Length - 1];

            for (int i = 0; i < addr.Length - 1; i++)
            {
                if(fullRange) rngResult[i] = getRange(getRow(addr[i]), 1, getRow(addr[i + 1]), maxColumns);
                else rngResult[i] = getRange(getRow(addr[i]), getColumn(addr[i]), getRow(addr[i + 1]), getColumn(addr[i + 1]));
            }

            return rngResult;
        }

        public Excel.Range transformRange(Excel.Range rng, int UpRow = 0, int DownRow = 0, int LeftCols = 0, int RightCols = 0)
        {
            string start = getRelativeAddress((rng.Cells[1, 1] as Excel.Range).Address, -UpRow, -LeftCols);
            string finish = getRelativeAddress((rng.Cells[rng.Rows.Count , rng.Columns.Count] as Excel.Range).Address, DownRow, LeftCols);

            return getRange(start, finish);
        }
        /*
        public Excel.Range transformRange(Excel.Range rng, int UpRow = 0, int DownRow = 0, int LeftCols = 0, int RightCols = 0)
        {
            string start, finish;

            if(UpRow == 0 && LeftCols == 0) return null;
            if((rng.Rows.Count+UpRow+DownRow) < 1 || (rng.Rows.Count+UpRow+DownRow) > maxRows) return null;
            if((rng.Columns.Count+LeftCols+RightCols) < 1 || (rng.Columns.Count+LeftCols+RightCols) > maxColumns) return null;

            start = (rng.Cells[1,1] as Excel.Range).Address;
            start = (wsExcel.Cells[getRow(start)-UpRow,getColumn(start)-LeftCols] as Excel.Range).Address;
            
            finish = (rng.Cells[rng.Rows.Count, rng.Columns.Count] as Excel.Range).Address;
            finish = (wsExcel.Cells[getRow(finish)+DownRow,getColumn(finish)+RightCols] as Excel.Range).Address;
            
            return (wsExcel.Range[start,finish] as Excel.Range);
            
        }
        */
        public Excel.Range getRange(int row1, int col1, int row2 = -1, int col2 = -1)
        {
            if (row2 > 0 && col2 > 0) return wsExcel.get_Range((wsExcel.Cells[row1, col1] as Excel.Range), (wsExcel.Cells[row2, col2] as Excel.Range));
            else return wsExcel.get_Range(wsExcel.Cells[row1, col1] as Excel.Range);
        }

        public Excel.Range getRange(string addr1, string addr2 = "NONE")
        {
            if (addr2 == "NONE") return wsExcel.get_Range((wsExcel.Range[addr1] as Excel.Range), (wsExcel.Range[addr2] as Excel.Range));
            else return wsExcel.get_Range(wsExcel.Range[addr1] as Excel.Range);
        }

        #endregion

        #region Data manipulation methods

        public string getCellValue(int rowInd, int colInd)
        {
            if (rowInd >= maxRows || colInd >= maxColumns || colInd < 1 || rowInd < 1) return null;
            return (wsExcel.Range[getAddress(rowInd, colInd)] as Excel.Range).Text;
            /*Excel.Range cellRange;
            string cellValue = "";

            cellRange = wsExcel.Cells[rowInd, colInd] as Excel.Range;
            //cellRange = wsExcel.get_Range(wsExcel.Cells[rowInd, colInd]);// as Excel.Range;
            
            if (cellRange.Text != null)
            {
                cellValue = Convert.ToString(cellRange.Text);
            }
            return cellValue;*/
        }

        public string getCellValue(string addr) { return (wsExcel.Range[addr] as Excel.Range).Text; }
        
        public string getCellValue(Excel.Range rng) { return rng.Text; }

        public string[] getRangeData(Excel.Range rng, bool Col = true, int Start = 0, int Finish = 0)
        {
            int cnt = Col ? rng.Rows.Count : rng.Columns.Count;
            
            if(Start < 0 || Finish < 0) return null;
            if(Start > 0) if(Start > cnt) return null;
            if (Finish > 0) if (Finish > cnt) return null;
            if ((Start > 0 && Finish > 0) && (Start > Finish)) return null;
            if (Start == 0) Start = 1;
            if (Finish == 0) Finish = cnt;

            string[] data = new string[Finish];
            
            for (int i = Start; i < Finish; i++)
            {
                data[i-1] = getCellValue((Col) ? i : 1, (Col) ? 1 : i);
            }
            return data;
        }

        public string[] getRangeData(string Address, bool Col = true, int Start = 0, int Finish = 0)
        {

            int cnt = Col ? maxRows : maxColumns;

            if (Start < 0 || Finish < 0) return null;
            if (Start > 0) if (Start > cnt) return null;
            if (Finish > 0) if (Finish > cnt) return null;
            if ((Start > 0 && Finish > 0) && (Start > Finish)) return null;
            if (Start == 0) Start = Col ? getRow(Address) : getColumn(Address);
            if (Finish == 0) Finish = cnt;

            string[] data = new string[Finish];
            int RC = Col ? getColumn(Address) : getRow(Address);


            for (int i = Start; i < Finish; i++)
            {
                data[i - 1] = getCellValue((Col) ? i : RC, (Col) ? RC : i);
            }
            return data;
        }

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

        public char getLetterCharacter(int number)
        {
            if (number < 1 || number > 26) return '-';
            char[] L = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
            return L[number - 1];
        }

        public int getLetterNumber(char C)
        {
            char[] L = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
            List<char> letters = new List<char>(L);
            C = Char.ToUpper(C);
            int index = letters.IndexOf(C);
            if (index >= 0) return index + 1;
            else return index;
        }

        #endregion

    }
}
