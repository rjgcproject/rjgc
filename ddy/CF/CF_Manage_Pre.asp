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
On Error Resume Next

If Session("CFCountUser")="" Then
 If Session("CFCountUser_View")="" Then
  Response.Cookies("CFCountUserCookie")=""
  Response.Cookies("CFCountUserCookie").Expires=Dateadd("n",-60,Now())
  Call AlertUrl("请重新登陆!","Index.asp")
 Else
  UserName=Session("CFCountUser_View")
 End If
Else
 UserName=Session("CFCountUser")
End If

Action=ChkStr(Request("Action"),1)

Tmp = HttpPath(2)
Ip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If Ip= "" Then Ip= Request.ServerVariables("REMOTE_ADDR")
%>


<%If Action="jsqset" Then
Set Rs= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
Rs.Open Sql,Conn,1,1
 If Rs("CounterShow")=-1 Then
  show=show&"ashow_1();"
 ElseIf Rs("CounterShow")=0 Then
  show=show&"ashow_2();"
 End If
 
 If Rs("ImgCounterShow")=-1 Then
  show=show&"bshow_1();"
 ElseIf Rs("ImgCounterShow")=0 Then
  show=show&"bshow_2();"
 End If

 If Rs("online")=-1 Then
  show=show&"cshow_1();"
 ElseIf Rs("online")=0 Then
  show=show&"cshow_2();"
 End If
 
 If Rs("TjOpen")=-1 Then
  show=show&"dshow_1();"
 ElseIf Rs("TjOpen")=0 Then
  show=show&"dshow_2();"
 End If
End If

%>