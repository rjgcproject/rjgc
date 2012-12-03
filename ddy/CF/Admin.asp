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
<%
If Session("CFCountAdmin")="ok" Then Response.Redirect "CF_Admin_Manage.asp"
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML>
<HEAD>
<TITLE><%=RsSet("SysTitle")%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#include file="CF_Style.asp"-->
<script language=JavaScript>

 function ccimgchange()
{
 document.getElementById("ccimg").src = 'CF_CheckCode.asp?a=' + Math.random();
}

function adminlogincheck()
{
  if(window.document.form1.admin.value=="")
  {
     alert("请输入管理员名称!");
	 window.document.form1.admin.focus();
	 return false;
  }
  if(window.document.form1.pwd.value=="")
  {
     alert("请输入管理员密码!");
	 window.document.form1.pwd.focus();
	 return false;
  }
 if(window.document.form1.checkcode.value=="")
  {
     alert("请输入四位数字的验证码!");
	 window.document.form1.checkcode.focus();
	 return false;
  }
  return true;
}

</script>
</HEAD>
<BODY>
<!--#include file="Top.asp"-->

<table class="tb_2">
          <tr class="tr_1"> 
            <td colspan="2">超级管理员入口</td>
          </tr>
          <form name="form1" method="post" action="?Action=adminloginsave" onSubmit="return adminlogincheck()">
            <tr> 
              <td width="44%" class="td_3">管理员名称：</td>
              <td><input name="admin" type="text" id="admin3" size="15"></td>
            </tr>
            <tr> 
              <td class="td_3">管理员密码：</td>
              <td><input name="pwd" type="password" id="pwd" size="15"></td>
            </tr>
            <tr> 
              <td class="td_3">Cookie：</td>
              <td><select name="cookies_time" id="cookies_time" style="width:115px;">
                  <option value="60" selected>不保留</option>
                  <option value="1440">保留一天</option>
                  <option value="43200">保留一个月</option>
                  <option value="5256000">永久保留</option>
                </select> </td>
            </tr>
              <TR>
                <TD class="td_3">验证码：</TD>
                <TD><INPUT id=checkcode type=text size=6 name=checkcode><a href="javascript:ccimgchange()" title="看不清楚？换一个！"><IMG id="ccimg" height=14 src="CF_CheckCode.asp" border="0"></a></TD>
              </TR>
            <tr> 
              <td></td> 
			  <td>  

                  <input type="submit" name="Submit" value="确定">
                  　 
                  <input type="reset" name="Submit523" value="取消">
&nbsp;&nbsp;&nbsp;&nbsp;<a href="Index.asp">返回首页</a>
                </td>
            </tr>
          </form>
        </table>   

<!--#include file="Bottom.asp"-->

</BODY></HTML>

<%Call ConnClose()%>