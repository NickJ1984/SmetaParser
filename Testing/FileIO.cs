using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;


namespace ConsoleApplication1
{
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
            */
        }
    }


    class FileIO_inDev  //копия FileIO
    {
        public string folder { get; private set; }
        public string[] files { get; private set; }
        public string[] filePath { get; private set; }
        public string[] ext { get; private set; }
        public DateTime[] fileCreationTime { get; private set; }
        public bool isExists { get; private set; }
        private int count;

        public FileIO_inDev() { }
        public FileIO_inDev(string path)
        {
            changeFolder(path);
        }

        public void changeFolder(string path)
        {
            if (Directory.Exists(path))
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
                files[i] = files[i].Substring(files[i].LastIndexOf('\\') + 1, files[i].Length - (files[i].LastIndexOf('\\') + 1 + ext[i].Length));
                //fileCreationTime[i] = File.GetCreationTime(filePath[i]); 
                fileCreationTime[i] = File.GetLastWriteTime(filePath[i]);
            }

            /*
            //Удаляем Thumb
            List<string> tmp = new List<string>(files);
            int index = Array.IndexOf(files, "Thumb");
            if (index >= 0) remove(index);
            */
        }
    }

}
