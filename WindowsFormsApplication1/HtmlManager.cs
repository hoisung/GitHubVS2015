using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml;

namespace WindowsFormsApplication1
{
    class HtmlManager
    {
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

        

        private string[] GetLinks(XmlDocument xml)
        {
            string[] links;
            int i = 0;
            int count = 0;
            XmlNodeList Xnlist = xml.GetElementsByTagName("a");
            count = Xnlist.Count;
            foreach (XmlNode xn in Xnlist)
            {
                if ((((XmlElement)xn).GetAttribute("id") == string.Empty) | (xn.InnerText == string.Empty))
                {
                    count--;
                }
            }

            links = new string[count];
            foreach (XmlNode xn in Xnlist)
            {
                if ((((XmlElement)xn).GetAttribute("id") != string.Empty) & (xn.InnerText != string.Empty))
                {
                    links.SetValue(((XmlElement)xn).GetAttribute("id"), i);
                    i++;
                }
            }
            for (i = 0; i < links.Length; i++)
            {
                links[i] = links[i].Replace("_", "$");
            }
            return links;
        }
    }
}
