using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
    class UnEncryptArgs
    {
        private int arg0Len, keylen;
        private string key1;
        private string key2;
        private string tempkey;
        public UnEncryptArgs(string key)
        {
            tempkey = key;
            keylen = key.Length;
            if (keylen == 0 | key == null | key == "") return;
            arg0Len = Convert.ToInt32(key.Substring(0, 2), 16);
            key1 = key.Substring(2, arg0Len);
            key2 = key.Substring(arg0Len + 2);
        }
        public string getkey1()
        {
            return key1;
        }
        public string getkey2()
        {
            return key2;
        }
        public bool check()
        {
            if (tempkey == "" | tempkey == null) return false;
            Regex rg = new Regex(@"^[A-F0-9]+$");
            bool isSuccess = rg.Match(tempkey).Success;
            if ((arg0Len == keylen - 34) & (isSuccess))
            { return true; }
            return false;
        }

    }
    class EncryptRES
    {
        //public string Key1;
        //public string Key2;
        public string keyt;
        public bool Success;
        private string message;
        public string Message
        {
            get { return message; }
            set { message = value; }
        }
        public EncryptRES()
        {
            // Key1 = "";
            // Key2 = "";
            keyt = "";
            Success = false;
            Message = "";
        }
    }
    class UnEncryptRES
    {
        public string Value;
        public bool Success;
        private string message;
        public string Message
        {
            get { return message; }
            set { message = value; }
        }
        public UnEncryptRES()
        {
            Value = "";
            Success = false;
            Message = "";
        }
    }
}
