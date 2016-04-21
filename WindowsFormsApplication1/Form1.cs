using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;

using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Linq;
namespace WindowsFormsApplication1
{
    
    public partial class Form1 : Form
    {
        string savePath="D:\\";
        int year, month, day;
        public Form1()
        {
            InitializeComponent();
            setdefaultUserInfo();
            formatDate();
            setDateBoxValue();
            textBox1.Text = "在“保存路径”中选择月报XLS文件保存路径.\r\n在‘开始日期’和‘结束日期’中设置需要获取的区域数据\r\n在ATMF框里‘atmf工号’中输入要输出月报的atmf工号\r\n*****************\r\n在自助框内‘自助工号’内写上自助工号和密码\r\n*****************\r\n*****************\r\n*****************\r\n***注意***更换部件列数据为错误数据，如果有更换备件，请自己修改！！！！！！！";
        }

        private void setdefaultUserInfo()
        {
            string version = Application.ProductVersion;
            INIOperater INIOper = new INIOperater(Application.StartupPath + "\\");
            if (INIOper.isFileExists())
            {
                if (version != INIOper.GetKey("infomation", "version"))
                {
                    File.Delete(Application.StartupPath + "\\" + "config.ini"); return;
                }
                UserId.Text = INIOper.GetKey("user", "atmfuser");
                //GGPassword.Text = INIOper.GetKey("user", "ggpass");
                GGuserId.Text = INIOper.GetKey("user", "gguserid");
                savePath = INIOper.GetKey("export","path");
                SavePathViewBox.Text = savePath = savePath == "" ? "d:\\" : savePath;
                string keyt = INIOper.GetKey("user", "keyt");
                //********************************************************
                if (keyt == "" | keyt == null) return;
                EncryptClass e = new EncryptClass();
                UnEncryptRES UERes = e.UnEncryptProcess(keyt);
                if (UERes.Success)
                    GGPassword.Text = UERes.Value.Replace(GGuserId.Text + " ", "");
                else
                {
                    MessageBox.Show(UERes.Message);
                }
            }
           
        }
        private void UpdateUserInfo(string keyt)
        {
            INIOperater INIOper = new INIOperater(Application.StartupPath + "\\");
            INIOper.WriteKey("user", "atmfuser", UserId.Text);
            INIOper.WriteKey("user", "gguserid", GGuserId.Text);
            INIOper.WriteKey("user", "keyt", keyt);
            //INIOper.WriteKey("user", "keyt2", keyt2);
            INIOper.WriteKey("export", "path", savePath);
            INIOper.WriteKey("infomation", "version",Application.ProductVersion);
        }
        public void formatDate()
        {
            DateTime nowdate = DateTime.Now;
            year = nowdate.Year;
            month = nowdate.Month;
            day = nowdate.Day;

        }
        private void setDateBoxValue()
        {
            int Tmonth = month,Tyear=year;
            if (month == 1)
            {
                Tmonth = 12;
                Tyear = year - 1;
            }
            else Tmonth = month - 1;
            StarttingDate.Text = new DateTime(Tyear,Tmonth,day).ToString("yyyy-MM-dd");
            EndingDate.Text = DateTime.Now.ToString("yyyy-MM-dd");
        }
       
      
     
       
    
        private void StarttingDate_TextChanged(object sender, EventArgs e)
        {
           // Regex RX = new Regex("(19|20)[0-9]{2}[- /.](0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])");
        }

        private void StarttingDate_Leave(object sender, EventArgs e)
        {
             Regex RX = new Regex("(19|20)[0-9]{2}[- /.](0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])");
            // StarttingDate.Text = RX.Match(StarttingDate.Text).Value;
            if (RX.IsMatch(StarttingDate.Text))
            {
                StarttingDate.Text = Regex.Replace(StarttingDate.Text, "[- /.]", "-");
            }
        }

        private bool checkUserID()
        {
            
            if (textBox1.Text == "")
                return false;
            return true;
        }
        private bool isHaveOtherDataFile()
        {
            return File.Exists("维修档案查询统计表.xls");
        }
       
        private void button4_Click(object sender, EventArgs e)
        {
            
            if (checkUserID())
            {
                SetupProcess setupP = new SetupProcess();
                AtmfProcess atmfP = new AtmfProcess();
                GgwowmProcess ggwowmP = new GgwowmProcess();
                XlsContentProcess xcp = new XlsContentProcess(savePath);

                StarttingDate.Text= DateTime.Parse( StarttingDate.Text).ToString("yyyy-MM-dd");
                EndingDate.Text = DateTime.Parse(EndingDate.Text).ToString("yyyy-MM-dd");
                byte[] atmfdata, ggwowmdata,setupdata;
                atmfdata = atmfP.Process(StarttingDate.Text.Trim(), EndingDate.Text.Trim(), UserId.Text.Trim());
                setupdata = setupP.Process(StarttingDate.Text.Trim(), EndingDate.Text.Trim(), UserId.Text.Trim());
                ggwowmdata = ggwowmP.process(GGuserId.Text.Trim(),GGPassword.Text.Trim(), StarttingDate.Text.Replace("-", ""),EndingDate.Text.Replace("-",""));
                if(ggwowmdata == null) { textBox1.Text = "自助工号不存在或密码输入不正确";return; }
                EncryptRES enc= doEncrypt(GGuserId.Text, GGPassword.Text);
                if (enc.Success) { UpdateUserInfo(enc.keyt); }
                string exportName= xcp.process(atmfdata, ggwowmdata,ggwowmP.station,setupdata,Application.StartupPath+"\\");
                if (exportName == "")
                {
                    textBox1.Text = "指定日期所获取工单数为零,无数据导出，请重新输入日期范围！";
                    MessageBox.Show("指定日期所获取工单数为零,无数据导出，请重新输入日期范围！");
                    return;
                }
                
                    textBox1.Text = "月报已生成.\r\n****************\r\n";
                textBox1.Text += string.Format("路  径: {0};\r\n\r\n文件名： {1};", savePath, exportName);
            }
            else
            { textBox1.Text = "工号为空，请输入工号!"; }
        }
        private EncryptRES doEncrypt(string part1, string part2)
        {
            string PassStr = part1 + " " + part2;
            EncryptClass encryptProc = new EncryptClass();
            EncryptRES ENCRes = encryptProc.EncryptProcess(PassStr);
            return ENCRes;
        }
        private void updatetextbox()
        {
            textBox1.Text = "正在获取月报中...";
        }

        private void EndingDate_Leave(object sender, EventArgs e)
        {
            Regex RX = new Regex("(19|20)[0-9]{2}[- /.](0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])");
            if (RX.IsMatch(EndingDate.Text))
            {
                EndingDate.Text = Regex.Replace(EndingDate.Text, "[- /.]", "-");
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            string path = "";
           if( folderBrowserDialog1.ShowDialog()==DialogResult.OK||folderBrowserDialog1.ShowDialog()==DialogResult.Yes)
                path =folderBrowserDialog1.SelectedPath;
            int length = path.Length;
            path =path.Substring(length-1)=="\\"?path:path+"\\";
            SavePathViewBox.Text = path;
            savePath = path;

        }

    }
}
