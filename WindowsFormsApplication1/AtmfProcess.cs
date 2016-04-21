using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Xml;

namespace WindowsFormsApplication1
{
    class AtmfProcess
    {
        WebClient webclient = new WebClient();
        XmlDocument XmlDoc = new XmlDocument();
        TATMPostData DiagInfoQuary = new TATMPostData();
        
        string CURRENT_COOKIE = "staffnum=37065; UserID=37065; lvl=4;";
        //string detailLink = "";
        string baseHost = "http://atmf.guoguang.com.cn:9812/";
        //******************date format*****************************
       // int year, month, day;

        //**************field of exchange data***********************
        //*****xlsx from online to the report file(xls)***********
       // XSSFWorkbook sourceWB;//源XLSX数据；
       // ReportFileColumns RFcolumns = new ReportFileColumns();
        //*******************END*************************************
        //DetailField DFList = new DetailField();

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
            webclient.Headers.Add("Host", "atmf.guoguang.com.cn:9812");
            //webclient.Headers.Add("Connection","Keep-Alive");
            webclient.Headers.Add("Cookie", CURRENT_COOKIE);
        }
        
        private void setdoPostHttpHeaders( string poststr, string referer)
        {
            setdoGetHttpHeaders();
            webclient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            webclient.Headers.Set("ContentLength", poststr.Length.ToString());
            webclient.Headers.Set("Referer", referer);

        }
        private void JumpToDownloadPage(string StarttingDate,string EndingDate,string UserId)
        {
            setdoGetHttpHeaders();
            byte[] pageData = webclient.DownloadData("http://atmf.guoguang.com.cn:9812/servicequary2.aspx");
            string pageHtml = Encoding.UTF8.GetString(pageData);
            XmlDoc = GetHtmlNodes(pageHtml);
            FormatingFormDatabase(XmlDoc);
            DiagInfoQuary.TextBox5 = StarttingDate;
            DiagInfoQuary.TextBox6 = EndingDate;
            DiagInfoQuary.tb_staffno = UserId;
        }
        public XmlDocument GetHtmlNodes(string pageHtml)
        {
            XmlDocument xml = new XmlDocument();
            string doctype = "<!DOCTYPE html[<!ENTITY nbsp \" \"> <!ELEMENT input ANY> <!ATTLIST input id ID #REQUIRED>  <!ELEMENT a ANY> <!ATTLIST a id ID #REQUIRED>]>";
            pageHtml = pageHtml.Substring(pageHtml.IndexOf("<html"));
            pageHtml = doctype + pageHtml;
            //pageHtml = pageHtml.Replace(" id", " ID");
            pageHtml = pageHtml.Replace("xmlns=\"http://www.w3.org/1999/xhtml\"", "");
            xml.Load(new StringReader(pageHtml));
            return xml;
            

        }
        private string getvalueById(XmlDocument DOC, string Id)
        {
            return DOC.GetElementById(Id).GetAttribute("value");
        }

        public void FormatingFormDatabase(XmlDocument xmlDOC)// GetFormdata(XmlNodeList XNL)
        {

            DiagInfoQuary.__VIEWSTATE = getvalueById(xmlDOC, "__VIEWSTATE").Replace("+", "%2B");
            DiagInfoQuary.__EVENTVALIDATION = getvalueById(xmlDOC, "__EVENTVALIDATION").Replace("+", "%2B");
            DiagInfoQuary.__VIEWSTATEGENERATOR = getvalueById(xmlDOC, "__VIEWSTATEGENERATOR").Replace("+", "%2B");

           
        }
 
        private string getPostArgs()
        {
            string tems = "";
            int i = 0;
            FieldInfo[] finfoes = DiagInfoQuary.GetType().GetFields();
            foreach (FieldInfo finfo in finfoes)
            {

                tems += finfo.Name + '=' + finfo.GetValue(DiagInfoQuary);
                if (i < finfoes.Length - 1)
                    tems += "&";
                i++;
            }
            return tems;
        }

        public byte[] doPost()
        {
            string postvalues = getPostArgs();
            postvalues = postvalues.Replace("&Button1=%E6%9F%A5++%E8%AF%A2", "&Button2=%E5%AF%BC++%E5%87%BA");
            string uri = "http://atmf.guoguang.com.cn:9812/servicequary2.aspx";
            setdoPostHttpHeaders(postvalues, uri);
            byte[] postData = Encoding.UTF8.GetBytes(postvalues);
            byte[] pageData = webclient.UploadData(uri, postData);
            return pageData;
        }

        private void SetupCookie(string userid)
        {
            CURRENT_COOKIE = string.Format("staffnum={0}; UserID={0}; lvl=4;", userid == "" ? "37067" : userid);
        }
        private byte[] doDownloadXLSX()
        {
            //GetDetail(links[0]);
            // doPost("detail");
           return doPost();
            //textBox1.Text = pageHtml;
        }


       public byte[] Process(string starttingDate,string EndingDate,string UserId)
        {               
                SetupCookie(UserId);
                JumpToDownloadPage(starttingDate,EndingDate,UserId);
                return  doDownloadXLSX();
        }

 /*       public string UrlDecode(string url)
        {
            return HttpUtility.UrlEncode(url);
        }*/
    }

}
