using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace WindowsFormsApplication1
{
    public struct RepairSearchStruct
    {
        /// <summary>
        /// 省
        /// </summary>
      //  string syhdq1 = "45";
       // string syhdq2 = "";
        /// <summary>
        /// 开始时间
        /// </summary>
        public string dksrq;
        /// <summary>
        /// 结束时间
        /// </summary>
        public string djsrq;

        /*string sbankid = "";
        string sbankcode = "";
        string sjqbh = "";
        string seventbh = "";
        string ssbmclb = "";
        string sgzdl = "";
        string sgzxl = "";
        string ssjlb = "";
        */
        /// <summary>
        /// 分部机构号
        /// </summary>
        public string syhjg;/*   default = "0206000000"(广西承包商);*/
        /// <summary>
        /// 工程师姓名
        /// </summary>
        public string sffry;
        /// <summary>
        /// 工程师工号
        /// </summary>
        public string sffryid;
        /* string dwxksrq = "";
         string dwxjsrq = "";
         string sbankmc = "";
         string sjqxh = "";
         */
    }
    class RepairSearchPostData
    {

       

        public string url = "http://gdxt.guoguang.com.cn:16789/ggwowms/RepairMr%3ERepairSearch.do";
        public string Referer = "http://gdxt.guoguang.com.cn:16789/ggwowms/RepairMr%3ERepairSearchMain.do";
        RepairSearchStruct postvalues;
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="dksrq">开始日期</param>
        /// <param name="djsrq">结束日期</param>
        /// <param name="sffry">工程师姓名</param>
        /// <param name="sffryid">工程师工号</param>
        public RepairSearchPostData(string Dksrq,string Djsrq,string Sffry, string Sffryid)
        {
            postvalues.syhjg = "0206000000";
            postvalues.dksrq = Dksrq;
            postvalues.djsrq = Djsrq;
            postvalues.sffry = Sffry;
            postvalues.sffryid = Sffryid;
        }
        public RepairSearchPostData(RepairSearchStruct postValues)
        {
            postvalues = postValues;
        }
        public string PostString
        {
            get
            {
                string Str= HttpUtility.UrlEncode( "syhdq1=45&syhdq2=&dksrq=" + postvalues.dksrq + "&djsrq=" + postvalues.djsrq + "&sbankid=&sbankcode=&sjqbh=&seventbh=" +
                 "&ssbmclb=&sgzdl=&sgzxl=&ssjlb=&syhjg=" + postvalues.syhjg + "&sffry=" + postvalues.sffry + "&sffryid=" + postvalues.sffryid + "&dwxksrq=&dwxjsrq=" +
                 "&sbankmc=&sjqxh=").Replace("%3d","=");
                Str = Str.Replace("%26", "&");
                return Str;
            }


        }
        public int PostStringLenght
        {
            get {
                return PostString.Length;
            }
        }


       
    }
}
