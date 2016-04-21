using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication1
{
    class LoginPostData
    {
       public string posturl ="http://gdxt.guoguang.com.cn:16789/ggwowms/security%3Elogin_check.do";
       public string postStr = "fid=nn0034&fpasswd=111111";
       public string referer = "http://gdxt.guoguang.com.cn:16789/ggwowms/login.jsp";
        public LoginPostData(string userid,string password)
        { postStr = "fid=" + userid + "&fpasswd=" + password; }
    }
}
