using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1
{
    class EncryptClass
    {
        
        public UnEncryptRES UnEncryptProcess(string SRCkey)
        {
            UnEncryptRES UnEvalue = new UnEncryptRES();
            UnEncryptArgs args = new UnEncryptArgs(SRCkey);
            if (!args.check())
            {
                 UnEvalue.Message = "密钥被非法改动，已失效！请重新输入密码";
                return UnEvalue;
            }
            byte[] source = getHexBytes(args.getkey1(), false);
            string mac = GetActivatedAdaptorMacAddress();
            if (mac == "") { UnEvalue.Message = "无可用网络！请在有网络环境下使用"; return UnEvalue; }
            byte[] keys = getHexBytes(GetActivatedAdaptorMacAddress(), true);
            byte[] md5Src = new byte[source.Length + keys.Length];

            keys.CopyTo(md5Src, 0);
            source.CopyTo(md5Src, keys.Length);
            //string value = "";
            if (MD5Encrypt(md5Src) == args.getkey2())
            {
                byte[] keytem = Encrypt(source, keys);

                UnEvalue.Value = Encoding.Default.GetString(keytem);
                UnEvalue.Success = true;
                return UnEvalue;
            }
            else { UnEvalue.Message = "检测到系统环境变动，密钥已失效！请重新输入密码！"; }
            return UnEvalue;


        }
        private string MD5Encrypt(byte[] source)
        {
            MD5 md5 = new MD5CryptoServiceProvider();

            byte[] t = md5.ComputeHash(source);
            StringBuilder sb = new StringBuilder(32);
            foreach (byte b in t)
            {
                sb.Append(b.ToString("X").PadLeft(2, '0'));
            }
            return sb.ToString();
        }
        public EncryptRES EncryptProcess(string SRCkey)
        {
            EncryptRES EncValue = new EncryptRES();
            if (SRCkey.Length == 0)
            {
                EncValue.Message = "输入字符串为空!";
                return EncValue;
            }
            byte[] sourcebytes = Encoding.Default.GetBytes(SRCkey);
            byte[] keysbytes = getHexBytes(GetActivatedAdaptorMacAddress(), true);
            byte[] reqbytes1 = Encrypt(sourcebytes, keysbytes);
            //key1 = reqbytes1;
            byte[] newtem = new byte[reqbytes1.Length + keysbytes.Length];
            //EncValue.Key1 =getStringbyByteGroup( newtem);
            keysbytes.CopyTo(newtem, 0);
            reqbytes1.CopyTo(newtem, keysbytes.Length);
            string MD5String = MD5Encrypt(newtem);
            EncValue.keyt= (reqbytes1.Length * 2).ToString("X2") + getStringbyByteGroup(reqbytes1) + MD5String; ; //Encoding.Default.GetString(sourcebytes);
            EncValue.Success = true;
            return  EncValue;
        }
        private string getStringbyByteGroup(byte[] value)
        {
            StringBuilder sb = new StringBuilder();
            foreach (byte tem in value)
            {
                sb.Append(tem.ToString("X2"));
            }
            return sb.ToString();
        }
        private byte[] Encrypt(byte[] sourcebytes, byte[] keysbytes)
        {

            int klen = keysbytes.Length;
            int len = sourcebytes.Length;
            byte[] values = new byte[len];
            for (int i = 0; i < len; i++)
            {
                values[i] = (byte)(sourcebytes[i] ^ keysbytes[i % klen]);
            }
            return values;
        }
        private string GetActivatedAdaptorMacAddress()
        {
            string mac = "";
            ManagementClass mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementObjectCollection moc = mc.GetInstances();
            foreach (ManagementObject mo in moc)
            {
                if (mo["IPEnabled"].ToString() == "True")
                {
                    mac = mo["MacAddress"].ToString();
                }
            }
            return mac;
        }
        private byte[] getHexBytes(string source, bool mac)
        {
            return getHexByteGroupbyStrings(getHexValues(source, mac));
        }
        private string[] getHexValues(string source, bool mac = false)
        {
            if (!mac)
                for (int i = 2; i < source.Length; i = i + 2)
                {
                    source = source.Insert(i, " ");
                    i++;
                }
            string[] hexvalues = mac ? (source.Split(':')) : (source.Split(' '));
            return hexvalues;
        }
        private byte[] getHexByteGroupbyStrings(string[] sources)
        {
            if (sources.Length == 0)
                return null;
            byte[] values = new byte[sources.Length];
            for (int i = 0; i < sources.Length; i++)
            {
                int temp = Convert.ToInt32(sources[i], 16);
                values[i] = Convert.ToByte(temp);
            }
            return values;
        }
    }
}
