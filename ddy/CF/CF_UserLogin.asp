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
<script language=JavaScript>
function ccimgchange()
{
 document.getElementById("ccimg").src = 'CF_CheckCode.asp?a=' + Math.random();
}

function userlogincheck()
{
  if(window.document.form1.username.value=="")
  {
     alert("请输入用户名!");
	 window.document.form1.username.focus();
	 return false;
  }
  if(window.document.form1.pwd.value=="")
  {
     alert("请输入密码!");
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

<TABLE class="tb_3" style="width:213px;margin-top:3px;margin-bottom:0px">
      <TR class="tr_1"> 
        <TD colspan="2" class="td_1">用户登录</TD>
      </TR>
	  
  <form name="form1" method="post" action="Index.asp?Action=userloginsave" onsubmit="return userlogincheck()">
<%
If Session("CFCountUser")="" Then
%>
      <TR> 
        <TD class="td_3" style="padding-top:5px">&nbsp;用户名：</TD>
        <TD style="padding-top:5px"><INPUT name=username id=username value="<%=ChkStr(Request("UserName"),1)%>" size=15></TD>
      </TR>
      <TR> 
        <TD class="td_3" style="padding-top:5px">&nbsp;密&nbsp;&nbsp;&nbsp;&nbsp;码：</TD>
        <TD style="padding-top:5px"><INPUT id=pwd type=password size=15 name=pwd></TD>
      </TR>
      <TR> 
        <TD class="td_3" style="padding-top:5px">&nbsp;Cookie：</TD>
        <TD style="padding-top:5px"><select name="cookies_time" id="cookies_time" style="width:115px;">
            <option value="60" selected>不保留</option>
            <option value="1440">保留一天</option>
            <option value="43200">保留一个月</option>
            <option value="5256000">永久保留</option>
          </select></TD>
      </TR>
      <TR> 
        <TD class="td_3" style="padding-top:3px">&nbsp;验证码：</TD>
        <TD style="padding-top:3px"><INPUT id=checkcode type=text size=6 name=checkcode><a href="javascript:ccimgchange()" title="看不清楚？换一个！"><IMG id="ccimg" height=14 src="CF_CheckCode.asp" border="0"></a></TD>
      </TR>
      <TR> 
        <TD class="td_3" style="padding-top:3px">&nbsp;<a href="PwdRecover.asp"><font color="#0066FF">密码找回</font></a></TD>
        <TD><INPUT type="submit" value="登&nbsp;录" name="submit" class="button">

 &nbsp;<a href="Reg.asp">注&nbsp;&nbsp;册</a></TD>
      </TR>
      <%Else%>
      <TR> 
        <TD colspan="2" style="text-align:center;height:142px;vertical-align:top">
		<INPUT onclick="location.href='Manage.asp';" type='button' value='进入管理' name='submit' class='button' style="margin-top:10px">
		<br />
		<INPUT onclick="location.href='Manage.asp?action=userlogout';" type='button' value='退出管理' name='submit' class='button'  style="margin-top:10px">
		</TD>
      </TR>

      <%End if%>
  </FORM>
</TABLE>
