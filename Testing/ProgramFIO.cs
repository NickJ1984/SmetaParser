/*using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace TestExcelIO
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application appExcel = new Excel.Application();
            Excel.Workbook wbExcel = appExcel.Workbooks.Open(@"C:\WORK\Log_2015-01-15_10-53-19.xlsx", ReadOnly:true);
            //Excel.Worksheet wsExcel = (Excel.Worksheet)wbExcel.Sheets[1];
            Excel.Worksheet wsExcel = (Excel.Worksheet)wbExcel.Worksheets.get_Item(1);
            Excel.Range rng;
            //Excel.Worksheet wsExcel = appExcel.ActiveSheet as Excel.Worksheet;
            object missingObj = System.Reflection.Missing.Value;
            int maxRows = wsExcel.UsedRange.Rows.Count;
            int maxColumns = wsExcel.UsedRange.Columns.Count;
            string cellValue = "";

            //Excel.Range cellRange = (Excel.Range)wsExcel.Cells[0, 0];
            //cellValue = wsExcel.UsedRange.Cells[0, 0].Value2;
            Excel.CellFormat cell;
            rng = wsExcel.Cells.Find("1111");
            cellValue = Convert.ToString(rng.Value2);
            Console.WriteLine(cellValue);
            Console.ReadLine();

            #region ExcelClose
            wbExcel.Close(false, missingObj, missingObj);
            appExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(appExcel);
            appExcel = null;
            wbExcel = null;
            wsExcel = null;
            System.GC.Collect();
            #endregion
        }
    }
}
*/


/*
class FileIO_
{
    private string folder;
    private string logExt = "*.xlsx";
    public string[] files { get; private set; }
    private string[] filePath;
       
    public FileIO_()
    {
        if(folder == "") folder = @"\\DB1CSQL\Smeta\Errors\";
        getFiles();
    }
    public FileIO_(string path) : base()
    {
        folder = path;
    }
        
    private void getFiles()
    {
        if (!Directory.Exists(folder)) return;
        filePath = Directory.GetFiles(folder);
        files = filePath;
        for (int i = 0; i < files.Length; i++) files[i] = files[i].Substring(files[i].LastIndexOf('\\')+1, files[i].Length - (files[i].LastIndexOf('\\') + 5));
            
        //Удаляем Thumb
        List<string> tmp = new List<string>(files);
        int index = Array.IndexOf(files, "Thumb");
        tmp.RemoveAt(index);
        files = tmp.ToArray();
    }
        
    public void outputFileList()
    {
        int cnt = 0;
        Console.WriteLine("File list in {0} directory:", folder);
        foreach (string s in files) Console.WriteLine("{0}: {1}", ++cnt, s);
            
    }
}
 */