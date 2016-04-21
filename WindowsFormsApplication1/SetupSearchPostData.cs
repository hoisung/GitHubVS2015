using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace WindowsFormsApplication1
{
    public struct SetupSearchStruct
    {
       public string __VIEWSTATE;
       public string __VIEWSTATEGENERATOR;
       public string __EVENTVALIDATION;
        /// <summary>
        /// 工号
        /// </summary>
        public string tb_setupresponse;
        /// <summary>
        /// 开始日期
        /// </summary>
        public string TextBox10;
        /// <summary>
        /// 结束日期
        /// </summary>
        public string TextBox11;
    }
    class SetupSearchPostData
    {



        public string url = "http://atmf.guoguang.com.cn:9812/setupquary.aspx";
        public string Referer = "http://atmf.guoguang.com.cn:9812/setupquary.aspx";
        SetupSearchStruct postvalues;
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="datefrom">开始日期</param>
        /// <param name="dateto">结束日期</param>
        /// <param name="ID">工程师工号</param>
        public SetupSearchPostData(string viewstate,string viewstategenerator,string eventvalidation, string datefrom, string dateto, string ID)
        {
            postvalues.__VIEWSTATE = viewstate;
            postvalues.__VIEWSTATEGENERATOR =viewstategenerator;
            postvalues.TextBox10 = datefrom;
            postvalues.TextBox11 = dateto;
            postvalues.tb_setupresponse= ID;
        }
        public SetupSearchPostData(SetupSearchStruct postValues)
        {
            postvalues = postValues;
        }
        public string PostString
        {
            get
            {
                string Str = HttpUtility.UrlEncode("__EVENTTARGET=&__EVENTARGUMENT=&__LASTFOCUS=&__VIEWSTATE="+postvalues.__VIEWSTATE+"&__VIEWSTATEGENERATOR="+postvalues.__VIEWSTATEGENERATOR+
                    "&__VIEWSTATEENCRYPTED=&__EVENTVALIDATION="+postvalues.__EVENTVALIDATION+"&TextBox1=&TextBox2=&tb_paiid=&tb_subbranchname=&tb_setupstaffno=&tb_setupresponse="+postvalues.tb_setupresponse+
                    "&tb_deviceno=&tb_banknum=&dd_province=广西&dd_city=全部&DropDownList5=全部&DropDownList6=全部&DropDownList1=全部&DropDownList3=全部&DropDownList7=全部&TextBox5=&TextBox6=&TextBox7=&TextBox8=&TextBox10="+
                    postvalues.TextBox10+"&TextBox11="+postvalues.TextBox11+ "&DropDownList8=全部&DropDownList4=全部&DropDownList2=全部&Button2=导  出").Replace("%3d", "=");
                Str = Str.Replace("%26", "&");
                Str = Str.Replace("%2f", "/");
                Str = Str.Replace("%2b", "%2B");
                return Str;
            }


        }
        public int PostStringLenght
        {
            get
            {
                return PostString.Length;
            }
        }
    }
    }
