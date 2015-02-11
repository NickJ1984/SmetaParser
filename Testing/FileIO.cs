using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;


namespace ConsoleApplication1
{
    /*
    class FileIO
    {
        public string folder { get; private set; }
        public string[] files { get; private set; }
        public string[] filePath { get; private set; }
        public string[] ext { get; private set; }
        public DateTime[] fileCreationTime { get; private set; }
        public bool isExists { get; private set; }
        private int count;

        public FileIO() { }
        public FileIO(string path)
        {
            changeFolder(path);
        }

        public void changeFolder(string path)
        {
            if(Directory.Exists(path))
            {
                folder = path;
                isExists = true;
            }
        }

        private void remove(int index)
        {
            if (index >= count) return;

            List<string> tFiles = new List<string>(files);
            List<string> tFilePath = new List<string>(filePath);
            List<string> tExt = new List<string>(ext);
            List<DateTime> tFileCreationTime = new List<DateTime>(fileCreationTime);
            
            tFiles.RemoveAt(index);
            tFilePath.RemoveAt(index);
            tExt.RemoveAt(index);
            tFileCreationTime.RemoveAt(index);

            files = tFiles.ToArray();
            filePath = tFilePath.ToArray();
            ext = tExt.ToArray();
            fileCreationTime = tFileCreationTime.ToArray();
        }

        public void getFiles(string srchPattern = "*")
        {
            if (!isExists) return;

            filePath = null;
            files = null;
            ext = null;
            fileCreationTime = null;

            System.GC.Collect();

            filePath = Directory.GetFiles(folder, srchPattern);
            files = new string[filePath.Length];
            ext = new string[filePath.Length];
            filePath.CopyTo(files, 0);
            filePath.CopyTo(ext, 0);
            fileCreationTime = new DateTime[files.Length];

            for (int i = 0; i < files.Length; i++)
            {
                ext[i] = ext[i].Substring(ext[i].LastIndexOf('.'), ext[i].Length - ext[i].LastIndexOf('.'));
                files[i] = files[i].Substring(files[i].LastIndexOf('\\') + 1, files[i].Length - (files[i].LastIndexOf('\\')+ 1 + ext[i].Length));
                //fileCreationTime[i] = File.GetCreationTime(filePath[i]); 
                fileCreationTime[i] = File.GetLastWriteTime(filePath[i]);
            }

            /*
            //Удаляем Thumb
            List<string> tmp = new List<string>(files);
            int index = Array.IndexOf(files, "Thumb");
            if (index >= 0) remove(index);
            
        }
    }*/ //Прошлая версия

    class FileIO  //переработан
    {
        private List<string> files = new List<string>();
        private List<string> fullpath = new List<string>();
        private List<string> ext = new List<string>();
        private List<DateTime> dateCreation = new List<DateTime>();
        public ust_LogFileDescription[] logfiles { get; private set;}

        public string folder { get; private set; }
        public string searchPattern { get; set; }
        
        public bool isExists { get; private set; }
        private bool isCollected;
        
        #region Constructors

        public FileIO() { }
        public FileIO(string path)
        {
            changeFolder(path);
        }

        #endregion

        #region Private

        private void remove(int index)
        {
            files.RemoveAt(index);
            fullpath.RemoveAt(index);
            ext.RemoveAt(index);
            dateCreation.RemoveAt(index);
        }

        private void clearFull()
        {
            files.Clear();
            fullpath.Clear();
            ext.Clear();
            dateCreation.Clear();
            folder = null;
            searchPattern = null;
            logfiles = null;

            isCollected = false;
            isExists = false;

            System.GC.Collect();
        }

        private void clearData()
        {
            files.Clear();
            fullpath.Clear();
            ext.Clear();
            dateCreation.Clear();
            logfiles = null;

            isCollected = false;

            System.GC.Collect();
        }

        private void getFilesList()
        {
            fullpath.AddRange(Directory.GetFiles(folder, searchPattern));
        }

        private void getData()
        {
            for (int i = 0; i < fullpath.Count; i++)
            {
                ext.Add(fullpath[i].Substring(fullpath[i].LastIndexOf('.'), fullpath[i].Length - fullpath[i].LastIndexOf('.')));
                files.Add(fullpath[i].Substring(fullpath[i].LastIndexOf('\\') + 1, fullpath[i].Length - (fullpath[i].LastIndexOf('\\') + 1 + ext[i].Length)));
                dateCreation.Add(File.GetLastWriteTime(fullpath[i]));
            }
        }

        private ust_LogFileDescription[] convertLFD()
        {
            ust_LogFileDescription[] lfd = new ust_LogFileDescription[fullpath.Count];

            for (int i = 0; i < lfd.Length; i++)
            {
                lfd[i].DateOfCreation = dateCreation[i];
                lfd[i].FileName = files[i];
                lfd[i].FullPath = fullpath[i];
            }

            return lfd;
        }

        #endregion

        #region Debug

        public void debug_printInfo(int index)
        {
            Console.WriteLine("Search pattern: {0}", searchPattern);
            Console.WriteLine("Path: {0}", folder);
            Console.WriteLine("Name: {0}", files[index]);
        }

        #endregion

        #region Public

        public void scan()
        {
            if(!isExists) return;
            clearData();

            getFilesList();
            getData();
            logfiles = convertLFD();
            isCollected = true;
        }

        public void changeFolder(string path)
        {
            if (Directory.Exists(path))
            {
                folder = path;
                isExists = true;
            }
        }

        public string[] getFiles()
        {   return files.ToArray(); }

        public string[] getExt()
        { return ext.ToArray(); }

        public string[] getFullPath()
        { return fullpath.ToArray(); }

        public DateTime[] getTime()
        { return dateCreation.ToArray(); }

        #endregion

    }

}
