using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.Serialization.Formatters.Binary;
using System.IO;

namespace ConsoleApplication1
{
    enum srzStatus { Normal = 0, Saved = 1, Loaded = 2, NotExist = -2, Error = -1, Null = -3 };
    class Serializer
    {
        public object obj;
        public srzStatus status { get; private set; }

        private Stream fstream;

        private string pth;
        public string path
        {
            set
            {
                status = (value == "" || value == null) ? srzStatus.Null : srzStatus.Normal;
                pth = value;
            }
            get
            {
                return pth;
            }
        }

        public Serializer()
        {
            path = "";
            status = srzStatus.Null;
        }

        public void Write()
        {
            BinaryFormatter bf = new BinaryFormatter();
            CreateStream();
            try
            { 
                bf.Serialize(fstream, obj);
                status = srzStatus.Saved;
            }
            catch { status = srzStatus.Error; }
            CloseStream();
        }

        public void Read()
        {
            BinaryFormatter bf = new BinaryFormatter();
            OpenStream();
            if (status == srzStatus.Error || status == srzStatus.NotExist) return;
            try
            {
                obj = bf.Deserialize(fstream);
                status = srzStatus.Loaded;
            }
            catch { status = srzStatus.Error; }
            CloseStream();
        }

        #region File stream methods

        private void CreateStream()
        {
            try
            { fstream = File.Create(path); }
            catch { status = srzStatus.Error; }
        }

        private void OpenStream()
        {
            if (!File.Exists(path))
            {
                status = srzStatus.NotExist;
                return;
            }
            try
            { fstream = File.Open(path, FileMode.Open); }
            catch { status = srzStatus.Error; }
        }

        private void CloseStream()
        {
            try
            { fstream.Close(); }
            catch { status = srzStatus.Error; }
        }

        #endregion

    }
}
