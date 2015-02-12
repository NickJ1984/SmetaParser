using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Runtime.Serialization.Json;

using StreamWriter = System.IO.StreamWriter;
using StreamReader = System.IO.StreamReader;

namespace ConsoleApplication1
{
    class JSONSerializer
    {
        DataContractJsonSerializer ser;
        StreamWriter sw;
        StreamReader sr;
        string fullpath;
        string filename;
        bool isExist;

        public JSONSerializer() { }

        

    }
}
