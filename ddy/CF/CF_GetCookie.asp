<%
'乘风多用户计数器
'作者QQ：178575
'作者EMail：yliangcf@163.com
'作者网站：http://www.qqcf.com
'详细简介：http://www.qqcf.com/?action=list&list=cfcount
'上面有程序在线演示，安装演示，使用疑难解答，最新版本下载等内容
'因为这些内容可能时常更新，就没有放在程序里，请自己上网站上查看
'有完整版本的演示
%>
<%
CFCountUserCookie=Chkstr(Request.Cookies("CFCountUserCookie"),1)

If Session("CFCountUser")="" Then'Session不存在但cookie存在时，重新获得Session
 If CFCountUserCookie<>"" Then
  Sql="Select UserName From CFCount_User Where UserCode='"&CFCountUserCookie&"'"
  Set Rs=Conn.Execute(Sql)
  If Not Rs.Eof Then
   Session("CFCountUser")=Rs("UserName")
  Else
   Response.Cookies("CFCountUserCookie")=""
   Response.Cookies("CFCountUserCookie").Expires=Dateadd("n",-60,Now())
  End If
  Rs.Close()
 End If
End If

If Session("CFCountUser")<>"" And CFCountUserCookie="" Then'Session存在但cookie不存在时，重新写Cookie
  Sql="Select UserCode From CFCount_User Where UserName='"&Session("CFCountUser")&"'"
  Set Rs=Conn.Execute(Sql)
  If Not Rs.Eof Then
   Response.Cookies("CFCountUserCookie")=Rs("UserCode")
   Response.Cookies("CFCountUserCookie").expires=Dateadd("n",60,Now())
  Else
   Session("CFCountUser")=""
  End If
  Rs.Close()
End If


CFCountAdminCookie=Chkstr(Request.Cookies("CFCountAdminCookie"),1)

If Session("CFCountAdmin")="" Then 
 If CFCountAdminCookie<>"" then   
  If RsSet("AdminCode")=CFCountAdminCookie Then Session("CFCountAdmin")="ok"
 End If
End If


If Session("CFCountAdmin")<>"" And CFCountAdminCookie="" Then'Session存在但cookie不存在时，重新写Cookie
   Response.Cookies("CFCountAdminCookie")=RsSet("AdminCode")
   Response.Cookies("CFCountAdminCookie").expires=Dateadd("n",60,Now())
End If
%>