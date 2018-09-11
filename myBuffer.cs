using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace shuzihua
{
    class myBuffer
    {
        //for word
        public List<string> value;
        public string section;
        public string item;
        public string item1;
        public List<string> id;
        public List<bool> valid;

        //for excel
        public Dictionary<string,Dictionary<string, string>> dic;
        public myBuffer()
        {
            value = new List<string>();
            section = "";
            item = "";
            item1 = "";
            id = new List<string>();
            valid = new List<bool>();
            dic = new Dictionary<string, Dictionary<string, string>>();
        }

    }
}
