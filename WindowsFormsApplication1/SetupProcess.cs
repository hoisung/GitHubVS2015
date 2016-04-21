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
    class SetupProcess
    {
        WebClient webclient = new WebClient();
        XmlDocument XmlDoc = new XmlDocument();
        SetupSearchPostData SetupPostData;
        SetupSearchStruct postdataStruct;
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

        private void setdoPostHttpHeaders(int poststrlenght, string referer)
        {
            setdoGetHttpHeaders();
            webclient.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            webclient.Headers.Set("ContentLength", poststrlenght.ToString());
            webclient.Headers.Set("Referer", referer);

        }
        private void JumpToDownloadPage(string StarttingDate, string EndingDate, string UserId)
        {
            setdoGetHttpHeaders();
            byte[] pageData = webclient.DownloadData("http://atmf.guoguang.com.cn:9812/setupquary.aspx");
            string pageHtml = Encoding.UTF8.GetString(pageData);
            XmlDoc = GetHtmlNodes(pageHtml);
            FormatingFormDatabase(XmlDoc);
            
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

            postdataStruct.__VIEWSTATE = getvalueById(xmlDOC, "__VIEWSTATE");//.Replace("+", "%2B");
            postdataStruct.__EVENTVALIDATION = getvalueById(xmlDOC, "__EVENTVALIDATION");//).Replace("+", "%2B");
            postdataStruct.__VIEWSTATEGENERATOR = getvalueById(xmlDOC, "__VIEWSTATEGENERATOR");//.Replace("+", "%2B");


        }


        public byte[] doPost()
        {
           
            setdoPostHttpHeaders(SetupPostData.PostStringLenght,SetupPostData.Referer);
            byte[] postData = Encoding.UTF8.GetBytes(SetupPostData.PostString);
            byte[] pageData = webclient.UploadData(SetupPostData.url, postData);
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


        public byte[] Process(string starttingDate, string EndingDate, string UserId)
        {
            SetupCookie(UserId);
            JumpToDownloadPage(starttingDate, EndingDate, UserId);
            postdataStruct.tb_setupresponse = UserId;
            postdataStruct.TextBox10 = starttingDate;
            postdataStruct.TextBox11 = EndingDate;
            SetupPostData = new SetupSearchPostData(postdataStruct);
            return doDownloadXLSX();
        }

        /*       public string UrlDecode(string url)
               {
                   return HttpUtility.UrlEncode(url);
               }*/
    }

}
