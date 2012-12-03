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
<!--#Include File="Conn.asp"-->
<!--#Include File="CF_MyFunCtion.asp"-->
<!--#Include File="CF_Md5.asp"-->
<!--#Include File="CF_Getcookie.asp"-->
<%
Action=Request("Action")
%>
<!--#Include File="CF_Do.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML><HEAD><TITLE><%=RsSet("SysTitle")%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#Include File="CF_Style.asp"-->
</HEAD>

<BODY>
<!--#Include File="Top.asp"-->

<table class="tb_2 tb_color" style="margin-top:4px;margin-bottom:0px">
  <tr>
    <td>&nbsp;&nbsp;&nbsp;<a href="Manage.asp?Action=lylist">来源明细</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=lytj">来源统计</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=onlinetj">在线统计</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=daytj">每日统计</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=hourtj">小时统计</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=backtj">回头率统计</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=searchtj">搜索引擎统计</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=keywordtj">搜索关键字统计</a>|&nbsp;&nbsp;<a href="Manage.asp?Action=jsqset">计数设置</a>
	</td>
  </tr>
</table>

<table align="center">
  <tr> 
    <td valign="top"><!--#Include File="CF_UserLogin.asp"--></td>
    <td><IMG src="images/banner.jpg" hspace="3" vspace="2" border="0"></td>
  </tr>
</table>

<TABLE class="tb_2" style="margin-top:2px;margin-bottom:0px;">
    <TR  class="tr_1"> 
      <TD  class="td_1"><%=RsSet("SysTitle")%>简介</TD>
    </TR>

    <TR> 
      <TD>

<TABLE width=100% border=0 align=center cellPadding=0 cellSpacing=0 
bgColor=#ffffff class="te2">
                    <TBODY>
                      <TR>
                        <TD vAlign=bottom align=middle width=10 height=218>&nbsp; </TD>
                        <TD width=706 >功　能： 
                          <br>
                          1.总共60种记数器图片样式可以选择，还有两种统计图标可以选择，完全满足你的需求。<br>
                          2.可以设置计数器显示数字，显示位数，计数器是否隐藏，统计信息是否公开等。<br>
                          3.页面显示记数和IP防刷新记数两种记数模式，支持Script网站方式和Img网店类方式调用计数器代码。<br>
                          4.可以记录来访客的来源IP地址和来源页面信息，在线人数。<br>
                          5.每月、每天和每小时的访问数据统计，回头率统计，每个网页浏览统计。<br>
                          6.搜索引擎统计，还可以自己定义搜索引擎，搜索关键字统计。<br>
                          7.注册用户找回密码功能，用户可以重置统计，删除注册账号。<br>
                          8.安全性：密码MD5加密，找回密码答案MD5加密，注册、登陆使用验证码，完全防Sql注入。<br> 特有的功能： 
                          <br>
                          1.核心使用缓存技术构建，内核十分强健。<br>
                          2.计数器数字图片和统计图标两种机制共存，众多设置可调。<br>
                          3.Script网站和Img网店类两种方式调用计数器，Img网店类方式计数器可以在任何能插入图片的地方使用。<br>
                          4.独有的错误自动修复机制，能在计数器发生错误后自动修复。<br>
                          5.完全杜绝并发线程容易对数据库造成的损坏，在流量大的网站上使用表现很稳定。<br>
                          6.缓存机制，在缓存中保存数据，操作常见动作，大量减少对数据库的增加，删除频繁的操作。<br>
                          7.稳定性、安全性、速度上表现都很优秀，功能齐全，代码集成程度高、完全公开，专业制作，完全免费。<br> 
                        </TD>
                        <TD vAlign=bottom align=middle width=21 height=218>&nbsp;</TD>
                      </TR>
                    </TBODY>
                  </TABLE>
				  
      </TD>
    </TR>
</TABLE>



<!--#Include File="Bottom.asp"-->
</BODY>
</HTML>

<%Call ConnClose()%>