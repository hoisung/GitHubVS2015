using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace WindowsFormsApplication1
{
    class GgwowmProcessRES { byte[] databuffer;bool Sucess=false; }
    class GgwowmProcess
    {
        WebClient webclient = new WebClient();
        String sessionStr;
        String host = "http://gdxt.guoguang.com.cn:16789";
        String UserID = "";
        String UserName = "";
        public string station = "";

        private void setdoGetHttpHeaders()
        {
            setdoGetHttpHeaders(false);
        }
        private void setdoGetHttpHeaders(Boolean withingCookies)
        {
            webclient.Headers.Clear();
            webclient.Headers.Add("Accept", "image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*");
            webclient.Headers.Add("Accept-Language", "zh-Hans-CN,zh-Hans;q=0.5");
            webclient.Headers.Add("Accept-Encoding", "gzip, deflate");
            webclient.Headers.Add("User-Agent", "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.3; Trident/8.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)");
            webclient.Headers.Add("Host", "gdxt.guoguang.com.cn:16789");
            if (withingCookies != false) { webclient.Headers.Add("Cookie", sessionStr); }
        }
        private void setdoPostHttpHeaders( string poststr, string referer)
        {
            setdoGetHttpHeaders();
            webclient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            webclient.Headers.Add("Cookie", sessionStr);
            webclient.Headers.Set("ContentLength", poststr.Length.ToString());
            webclient.Headers.Set("Referer", referer);

        }
        private byte[] Login(string userid,string password)
        {
            LoginPostData postdata = new LoginPostData(userid, password);
            byte[] pagedata = doPost(postdata.posturl, postdata.postStr, postdata.referer);
            string temp = Encoding.UTF8.GetString(pagedata);
            Regex rgx = new Regex("(?<=该用户不存在或用户及密码输入不正确)");
            if (rgx.Match(temp).Success) return null;
            return pagedata;
        }
        public byte[] process(string userid,string password,string dksrq,string djsrq)
        {
            if (sessionStr == null | sessionStr == "") getSession();
            /* string posturl = host + "/ggwowms/security%3Elogin_check.do";
             string postStr = "fid=nn0034&fpasswd=111111";
             string referer = "http://gdxt.guoguang.com.cn:16789/ggwowms/login.jsp";
 */
            RepairSearchStruct RSPostValues = new RepairSearchStruct();
            UserID = userid;
            byte[] pagedata;
            pagedata = Login(userid,password);
            if (pagedata == null) return null;
            getuserinfo();
            RSPostValues.sffry = UserName;
            RSPostValues.sffryid = userid;
            RSPostValues.dksrq = dksrq;
            RSPostValues.djsrq = djsrq;
            pagedata = gotorepairSearchMainpage();
            pagedata = getRepairListpagedata(RSPostValues);
            //textBox1.Text = Encoding.UTF8.GetString(pagedata);
            pagedata = getExportXlsData(getExportPostArgs(pagedata));
            return pagedata;
           // FileStream fs = new FileStream(Application.StartupPath + "\\bc.xls", FileMode.Create);
           // fs.Write(pagedata, 0, pagedata.Length);
           // fs.Close();
        }
        private ExportXlsStruct getExportPostArgs(byte[] pagedata)
        {
            //  XmlDocument xd = GetHtmlNodes(Encoding.UTF8.GetString(pagedata));
            string pagehtml = Encoding.UTF8.GetString(pagedata);
            ExportXlsStruct EXStruct;
            // EXStruct.datacount = getvalueById(xd, "datacount");
            // EXStruct.sqlForCount = getvalueById(xd, "sqlForCount");
            // EXStruct.sqlForList = getvalueById(xd, "sqlForList");
            EXStruct.datacount = getValuebyRegex("(?<=<input type=\"hidden\" name=\"datacount\" value=\").*(?=\" />)", pagehtml);
            EXStruct.sqlForCount = getValuebyRegex("(?<=<input type=\"hidden\" name=\"sqlForCount\" value=\").*(?=\" />)",pagehtml);
            EXStruct.sqlForList = getValuebyRegex("(?<=<input type=\"hidden\" name=\"sqlForList\" value=\").*(?=\" />)", pagehtml);
            return EXStruct;
        }
        private byte[] getRepairListpagedata(RepairSearchStruct RSPostValues)
        {
            RepairSearchPostData postdata = new RepairSearchPostData(RSPostValues.dksrq,RSPostValues.djsrq,RSPostValues.sffry,RSPostValues.sffryid);
            //RepairSearchPostData postdata = new RepairSearchPostData("20151001", "20151115", "高海淞", "nn0034");
            return doPost(postdata.url, postdata.PostString, postdata.Referer);
        }
        private byte[] getExportXlsData(ExportXlsStruct EXPostValues)
        {
            /*ExportXlsPostString postdata = new ExportXlsPostString("14",
                "select count (*) from (select a.*,rownum from (select a.*,b.spfry,b.dpfsj,b.dyyddsj,b.dxysj,b.dddsj,b.dwwcsj,b.dxcwcsj,b.ddallsj,b.dgbsj, to_char(a.dbxsj,'yyyy-mm-dd hh24:mi') bxsj,round((b.dxcwcsj-b.dddsj)*24*60,0) costtime, to_char(b.dpfsj,'yyyy-mm-dd hh24:mi') pfsj,to_char(b.dxysj,'yyyy-mm-dd hh24:mi') xysj,to_char(b.dddsj,'yyyy-mm-dd hh24:mi') ddsj,to_char(b.DXCWCSJ,'yyyy-mm-dd hh24:mi') XCWCSJ,to_char(b.DWWCSJ,'yyyy-mm-dd hh24:mi') DWCSJ,to_char(b.DDALLSJ,'yyyy-mm-dd hh24:mi') DALLSJ,decode(b.shfbz,'0','未回访','1','已回访') hfbzmc,c.szzbh,decode(a.ssjzt,'8',a.sfzpmx,c.sxcgzms) sxcgzms,c.scljg,c.nwxsl,decode(c.sfffs,'1','现场服务','2','电话服务') fffsmc,decode(c.skhpj,'1','满意','2','一般','不满意') khpjmc, (select sname from f_code t where (t.scatid='15') and (t.sid=a.ssjlb) ) sjlbmc,(select sname from f_code t where (t.scatid='10') and (t.sid=a.ssbmclb) ) sbmclbmc,(select spad from f_code t where (t.scatid='3') and (t.sid=a.syhdq)) yhdqmc ,(select sbankname from f_bank t where (t.sbankid=a.sbankid)) bankmc ,(select sbankname from f_bank t where t.slevel='1' and (t.sbankid=substr(a.sbankid,1,3)||'000000')) yhhymc, (select decode(slevel,'1','总行','2','分行','3','支行','4','网点') from f_bank where sbankid=a.sbankid) levelmc, (select sgzmc from f_bugtype t where (t.sgzlb=a.sgzdl))||'-'||(select sgzmc from f_bugtype t where (t.sgzlb=a.sgzxl)) gzlbmc,(select sname from f_code t where (t.scatid='7') and (t.sid=a.ssjzt)) sjztmc,(select sname from f_branch t where t.sid=a.syhjg) yhjgmc,(select syhxm from f_oper t where t.syhdm=a.sffry) ffrymc  from z_event a,z_event_do b,z_repaire c where a.seventbh=b.seventbh and a.seventbh=c.seventbh(+) and a.ssjzt in ('7','0','8') and a.syhdq like '45%'   and to_char(a.dbxsj,'yyyymmdd')>='20151001'  and to_char(a.dbxsj,'yyyymmdd')<='20151115' and a.syhjg='0206000000'  and a.sffry='nn0034' order by a.dczrq) a)"
                , "select a.*,rownum from (select a.*,b.spfry,b.dpfsj,b.dyyddsj,b.dxysj,b.dddsj,b.dwwcsj,b.dxcwcsj,b.ddallsj,b.dgbsj, to_char(a.dbxsj,'yyyy-mm-dd hh24:mi') bxsj,round((b.dxcwcsj-b.dddsj)*24*60,0) costtime, to_char(b.dpfsj,'yyyy-mm-dd hh24:mi') pfsj,to_char(b.dxysj,'yyyy-mm-dd hh24:mi') xysj,to_char(b.dddsj,'yyyy-mm-dd hh24:mi') ddsj,to_char(b.DXCWCSJ,'yyyy-mm-dd hh24:mi') XCWCSJ,to_char(b.DWWCSJ,'yyyy-mm-dd hh24:mi') DWCSJ,to_char(b.DDALLSJ,'yyyy-mm-dd hh24:mi') DALLSJ,decode(b.shfbz,'0','未回访','1','已回访') hfbzmc,c.szzbh,decode(a.ssjzt,'8',a.sfzpmx,c.sxcgzms) sxcgzms,c.scljg,c.nwxsl,decode(c.sfffs,'1','现场服务','2','电话服务') fffsmc,decode(c.skhpj,'1','满意','2','一般','不满意') khpjmc, (select sname from f_code t where (t.scatid='15') and (t.sid=a.ssjlb) ) sjlbmc,(select sname from f_code t where (t.scatid='10') and (t.sid=a.ssbmclb) ) sbmclbmc,(select spad from f_code t where (t.scatid='3') and (t.sid=a.syhdq)) yhdqmc ,(select sbankname from f_bank t where (t.sbankid=a.sbankid)) bankmc ,(select sbankname from f_bank t where t.slevel='1' and (t.sbankid=substr(a.sbankid,1,3)||'000000')) yhhymc, (select decode(slevel,'1','总行','2','分行','3','支行','4','网点') from f_bank where sbankid=a.sbankid) levelmc, (select sgzmc from f_bugtype t where (t.sgzlb=a.sgzdl))||'-'||(select sgzmc from f_bugtype t where (t.sgzlb=a.sgzxl)) gzlbmc,(select sname from f_code t where (t.scatid='7') and (t.sid=a.ssjzt)) sjztmc,(select sname from f_branch t where t.sid=a.syhjg) yhjgmc,(select syhxm from f_oper t where t.syhdm=a.sffry) ffrymc  from z_event a,z_event_do b,z_repaire c where a.seventbh=b.seventbh and a.seventbh=c.seventbh(+) and a.ssjzt in ('7','0','8') and a.syhdq like '45%'   and to_char(a.dbxsj,'yyyymmdd')>='20151001'  and to_char(a.dbxsj,'yyyymmdd')<='20151115'       and a.syhjg='0206000000'  and a.sffry='nn0034'   order by a.dczrq) a");
    */
            ExportXlsPostString postdata = new ExportXlsPostString(EXPostValues.datacount, EXPostValues.sqlForCount, EXPostValues.sqlForList);
            return doPost(postdata.url, postdata.PostString, postdata.Referer);
        }

        public void getuserinfo()
        {
            setdoGetHttpHeaders(true);
            webclient.Headers.Add("referer", "http://gdxt.guoguang.com.cn:16789/ggwowms/index.jsp");
            byte[] pageData = webclient.DownloadData("http://gdxt.guoguang.com.cn:16789/ggwowms/Mfw%3Ehomepage.do");
            string pagehtml = Encoding.UTF8.GetString(pageData);
            string retexStr = @"(?<="+UserID + @"\().*(?=\))";
            Regex RX = new Regex(retexStr);
            UserName = RX.Match(pagehtml).Value;
            retexStr = @"(?<=\d{10}\().*(?=服务站\))";
            RX = new Regex(retexStr);
            station = RX.Match(pagehtml).Value;

        }
        private byte[] gotorepairSearchMainpage()
        {
            setdoGetHttpHeaders(true);
            byte[] pageData = webclient.DownloadData("http://gdxt.guoguang.com.cn:16789/ggwowms/RepairMr%3ERepairSearchMain.do");
            return pageData;
        }
        private void getSession()
        {
            setdoGetHttpHeaders();
            byte[] pageData = webclient.DownloadData("http://gdxt.guoguang.com.cn:16789/ggwowms/login.jsp");
            sessionStr = webclient.ResponseHeaders.Get("Set-Cookie");
            if (sessionStr != null) sessionStr = sessionStr.Replace("; Path=/ggwowms", "");
        }
        private byte[] doPost(string posturl, string poststr, string referer)
        {
            setdoPostHttpHeaders(poststr, referer);
            byte[] pagedata = webclient.UploadData(posturl, Encoding.UTF8.GetBytes(poststr));
            return pagedata;
        }
        public string getValuebyRegex(string pattern,string pagehtml)
        {
            Regex RX = new Regex(pattern);
            string value = RX.Match(pagehtml).Value;
            return value;
        }
        public XmlDocument GetHtmlNodes(string pageHtml)
        {
            XmlDocument xml = new XmlDocument();
            string doctype = "<!DOCTYPE html[<!ENTITY nbsp \" \"> <!ELEMENT input ANY> <!ATTLIST input id ID #REQUIRED>  <!ELEMENT a ANY> <!ATTLIST a id ID #REQUIRED>]>";
            pageHtml = pageHtml.Substring(pageHtml.IndexOf("<body"));
            pageHtml = doctype+ "<html>" + pageHtml;
            //pageHtml = pageHtml.Replace(" id", " ID");
            pageHtml = pageHtml.Replace("xmlns=\"http://www.w3.org/1999/xhtml\"", "");
            pageHtml = pageHtml.Replace("name=", "ID=");
            xml.Load(new StringReader(pageHtml));
            return xml;


        }
        private string getvalueById(XmlDocument DOC, string Id)
        {
            return DOC.GetElementById(Id).GetAttribute("value");
        }
    }
}
