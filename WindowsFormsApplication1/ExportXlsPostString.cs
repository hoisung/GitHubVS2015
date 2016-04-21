using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace WindowsFormsApplication1
{
    public struct ExportXlsStruct
    {
       public string datacount;
       public string sqlForCount;
       public string sqlForList;
    } 
    class ExportXlsPostString
    {
        ExportXlsStruct postvalues;
        public string url = "http://gdxt.guoguang.com.cn:16789/ggwowms/RepairMr%3ERepairSearchExport.do?expbz=all";
        public string Referer = "http://gdxt.guoguang.com.cn:16789/ggwowms/RepairMr%3ERepairSearch.do";
        public ExportXlsPostString(string Datacount,string sqlforcount,string sqlforlist)
        {
            postvalues.datacount = Datacount;
            postvalues.sqlForCount = sqlforcount;
            postvalues.sqlForList = sqlforlist;
        }
        public ExportXlsPostString(ExportXlsStruct Postvalues)
        {
            postvalues = Postvalues;
        }
        public string PostString
        {
            get {
                string Str = HttpUtility.UrlEncode("datacount="+ postvalues.datacount +"&sqlForCount="+ postvalues.sqlForCount +"&sqlForList="+ postvalues.sqlForList).Replace("%3d","=");
                Str = Str.Replace("%26", "&");
                Str = Str.Replace("(", "%28");
                Str = Str.Replace(")", "%29");
                return Str;
            }
        }
        public int PostStringLenght
        {
            get { return PostString.Length; }
        }
    }
}
