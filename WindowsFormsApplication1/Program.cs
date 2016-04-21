using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            //*******处理DLL资源（如不存在则释放DLL，存在则跳过******
            string path = Application.StartupPath + "\\";
            string[] dllFileName = { "NPOI" ,"NPOI.OOXML","NPOI.OpenXml4Net", "ICSharpCode.SharpZipLib" , "NPOI.OpenXmlFormats" };
            foreach (string dllfile in dllFileName)
            {
                if (!File.Exists(path + dllfile+".dll"))
                {
                    FileStream fs = new FileStream(path + dllfile+".dll", FileMode.CreateNew, FileAccess.Write);
                    byte[] bytes= new byte[1024*1024*2];
                    if (dllfile == "NPOI")
                    {
                        bytes = Properties.Resources.NPOI;
                    }
                    else if (dllfile == "NPOI.OOXML")
                    {
                        bytes = Properties.Resources.NPOI_OOXML;
                    }
                    else if (dllfile == "NPOI.OpenXml4Net")
                    {
                        bytes = Properties.Resources.NPOI_OpenXml4Net;
                    }
                    else if (dllfile == "ICSharpCode.SharpZipLib")
                    {
                        bytes = Properties.Resources.ICSharpCode_SharpZipLib;
                    }
                    else if (dllfile == "NPOI.OpenXmlFormats")
                    {
                        bytes = Properties.Resources.NPOI_OpenXmlFormats;
                    }
                    
                    fs.Write(bytes, 0, bytes.Length);
                    fs.Close();
                } }
            //****************************************************
            //****************************************************
            Application.Run(new Form1());
        }
    }
}
