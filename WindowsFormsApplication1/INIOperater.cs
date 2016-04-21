using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace WindowsFormsApplication1
{
    class INIOperater
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filepath);
        [DllImport("kernel32")]
        private static extern long GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filePath);


        private string strFilePath = "config.ini";
        private string strSec = "";

        public INIOperater(string path)
        {
            strFilePath = path + strFilePath;
        }
        public bool isFileExists()
        {
            return File.Exists(strFilePath);
        }
        public object WriteKey(string sectionName,string key,string value)
        {
            try
            {
                WritePrivateProfileString(sectionName, key, value, strFilePath);
            }
            catch (Exception e)
            {
                return e.Message.ToString();
            }
            return true;
        }
        public string GetKey(string Section, string key)
        {
            StringBuilder SB = new StringBuilder(1024);
            GetPrivateProfileString(Section, key, "", SB, 1024, strFilePath);
            return SB.ToString();
        }
        public int GetNumKey(string Section, string key)
        {
            string tem = GetKey(Section, key);
            if (tem == "")
                return 000;
            return int.Parse(tem);
        }
    }
}
