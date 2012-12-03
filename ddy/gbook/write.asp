<!--#include file="conn.asp"-->
<%
if request("send")="ok" then

	username=trim(request.form("username"))
	usermail=trim(request.form("usermail"))

	if username="" or request.form("Comments")="" then
	response.write "<script language='javascript'>"
	response.write "alert('填写资料不完整，请检查后重新输入！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if checktxt(request.form("username"))<>request.form("username") then
	response.write "<script language='javascript'>"
	response.write "alert('您输入的用户名中含有非法字符，请检查后重新输入！');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
	end if

	if mailyes=0 then		'邮箱为必填时检查邮箱是否合法

	if checktxt(request.form("usermail"))<>request.form("usermail") then
	response.write "<script language='javascript'>"
	response.write "alert('您输入的邮箱中含有非法字符，请检查后重新输入！');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
	end if

	if Instr(usermail,".")<=0 or Instr(usermail,"@")<=0 or len(usermail)<10 or len(usermail)>50 then
	response.write "<script language='javascript'>"
	response.write "alert('您输入的电子邮件地址格式不正确，请检查后重新输入！');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
	end if

	end if

	if len(request.form("Comments"))>maxlength then
	response.write "<script language='javascript'>"
	response.write "alert('留言内容太长了，请不要超过"&maxlength&"个字符！');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	Comments1=request.form("Comments")
	bad1=split(bad,"/")		'过滤脏话
	for t=0 to ubound(bad1)
	Comments1=replace(Comments1,bad1(t),"***")
	next

	if request.form("Comments")<>Comments1 then
	response.write "<script language='javascript'>"
	response.write "alert('出错了，您的留言包含禁止提交的内容！');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
	end if


	set rs=Server.CreateObject("ADODB.RecordSet")
	sql="select * from Feedback where online='1' order by Postdate desc"
	rs.open sql,conn,1,3

			rs.Addnew
			rs("username")=Request("username")
			rs("comments")=Request("comments")
			rs("usermail")=Request("usermail")
			rs("face")=Request("face")
			rs("pic")=Request("pic")
			rs("url")=Request("url")
			rs("qq")=Request("qq")
			view=cstr(view)
			if view<>"0" then view="1"
			rs("online")=view
			rs("IP")=Request.serverVariables("REMOTE_ADDR")
			rs.Update
		rs.close
		set rs=nothing
	response.write "<script language='javascript'>"
	if view="0" then
	response.write "alert('留言提交成功，留言须经管理员审核才能发布。');"
	else
	response.write "alert('留言提交成功，单击“确定”返回留言列表！');"
	end if
	response.write "location.href='index.asp';"	
	response.write "</script>"
	response.end
end if
%>



<HTML><HEAD>
<TITLE><%=sitename%></TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script src="flashin.js" type="text/javascript"></script>
<meta name="description" content="<%=sitename%>">
<meta name="keywords" content="<%=sitename%>">
<link href="../css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style8 {font-size: 12px;
	color: #000000;
}
body,td,th {
	font-size: 12px;
}
.style14 {color: #333333;
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-weight: normal;
	text-decoration: none;
}
.style14 {font-size: 12px;
	color: #333333;
	font-family: Arial, Helvetica, sans-serif;
	line-height: 25px;
	font-weight: normal;
	text-decoration: none;
}
.style19 {font-size: 13px;
	color: #FFFFFF;
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style19 {font-size: 13px;
	color: #FFFFFF;
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style20 {font-size: 12px;
	font-family: Arial, Helvetica, sans-serif;
}
.style9 {font-size: 12px;
	color: #333333;
	font-family: Arial, Helvetica, sans-serif;
	line-height: 20px;
	font-weight: normal;
	text-decoration: none;
}
-->
</style>
<link href="../image/fish.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style17 {color: #4D679B;
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-weight: normal;
	text-decoration: none;
}
.style24 {font-size: 12px;
	font-family: Arial, Helvetica, sans-serif;
	color: #333333;
	text-decoration: none;
}
a:link {
	color: #333333;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #333333;
}
a:hover {
	text-decoration: underline;
	color: #333333;
}
a:active {
	text-decoration: none;
	color: #333333;
}
.style12 {font-size: 12px;
	color: #FFFFFF;
	font-family: Arial, Helvetica, sans-serif;
	line-height: 18px;
	font-weight: normal;
	text-decoration: none;
}
-->
</style>
</HEAD>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-top:10px">
  <tr>
    <td><table width="100%" border=0 align=center cellpadding=0 cellspacing=0 bgcolor="#FFFFFF" class="style14">
      <tr>
        <td>
          <form action=write.asp method=post name="book">
            <table width="99%" border="0" align="center" cellPadding="4" cellSpacing="1" borderColorLight="#000000" borderColorDark="#ffffff" class="style20">
              <tr bgColor="#ebebeb">
                <td   width="20%" align=right class="style24">您的姓名：</td>
                <td height="25"><input type=text name="UserName" size="30" maxlength=16>
                    <font color="#FF0000">*</font></td>
              </tr>
              <tr bgColor="#ebebeb">
                <td   width="20%" align=right class="style24">您的邮箱：</td>
                <td height="25" ><input type=text name="UserMail" size="30"  maxlength=50>
                    <%if mailyes=0 then%>
                    <font color="#FF0000">*</font>
                    <%end if%></td>
              </tr>
              
              <tr bgColor="#ebebeb">
                <td   width="20%" align=right class="style24">其它联系方式：</td>
                <td height="25"><input type=text value="" name="QQ" size="30"  maxlength=50>
                  <span class="style24">              （如QQ）</span></td>
              </tr>
              <tr bgColor="#ebebeb">
                <td   width="20%" align=right bgcolor="#ebebeb"><span class="style24">内容：</span><br>
                    <font color=red>（<%=maxlength%>字以内）</font></td>
                <td bgcolor="#ebebeb"><textarea name="Comments" rows="7" cols="66" style="overflow:auto;"></textarea></td>
              </tr>
              <tr bgColor="#ebebeb">
                <td   width="20%" align=right class="style24">请选择表情：</td>
                <td>
                  <input type="radio" value="1" name="face" checked>
                  <img border=0 src="images/face/face1.gif">
                  <input type="radio" value="2" name="face">
                  <img border=0 src="images/face/face2.gif">
                  <input type="radio" value="3" name="face">
                  <img border=0 src="images/face/face3.gif">
                  <input type="radio" value="4" name="face">
                  <img border=0 src="images/face/face4.gif">
                  <input type="radio" value="5" name="face">
                  <img border=0 src="images/face/face5.gif">
                  <input type="radio" value="6" name="face">
                  <img border=0 src="images/face/face6.gif">
                  <input type="radio" value="7" name="face">
                  <img border=0 src="images/face/face7.gif">
                  <input type="radio" value="8" name="face">
                  <img border=0 src="images/face/face8.gif">
                  <input type="radio" value="9" name="face">
                  <img border=0 src="images/face/face9.gif">
                  <input type="radio" value="10" name="face">
                  <img border=0 src="images/face/face10.gif"> <br>
                  <input type="radio" value="11" name="face">
                  <img border=0 src="images/face/face11.gif">
                  <input type="radio" value="12" name="face">
                  <img border=0 src="images/face/face12.gif">
                  <input type="radio" value="13" name="face">
                  <img border=0 src="images/face/face13.gif">
                  <input type="radio" value="14" name="face">
                  <img border=0 src="images/face/face14.gif">
                  <input type="radio" value="15" name="face">
                  <img border=0 src="images/face/face15.gif">
                  <input type="radio" value="16" name="face">
                  <img border=0 src="images/face/face16.gif">
                  <input type="radio" value="17" name="face">
                  <img border=0 src="images/face/face17.gif">
                  <input type="radio" value="18" name="face">
                  <img border=0 src="images/face/face18.gif">
                  <input type="radio" value="19" name="face">
                  <img border=0 src="images/face/face19.gif">
                  <input type="radio" value="20" name="face">
                  <img border=0 src="images/face/face20.gif"> </td>
              </tr>
              <tr bgColor="#ebebeb">
                <td width="20%" align=right bgcolor="#ebebeb" class="style24">请选择头像：</td>
                <td bgcolor="#ebebeb">
                  <input type="radio" value="1" name="pic" checked>
                  <img border=0 src="images/face/pic1.gif" width=60>
                  <input type="radio" value="2" name="pic">
                  <img border=0 src="images/face/pic2.gif" width=60>
                  <input type="radio" value="3" name="pic">
                  <img border=0 src="images/face/pic3.gif" width=60>
                  <input type="radio" value="4" name="pic">
                  <img border=0 src="images/face/pic4.gif" width=60>
                  <input type="radio" value="5" name="pic">
                  <img border=0 src="images/face/pic5.gif" width=60> <br>
                  <input type="radio" value="6" name="pic">
                  <img border=0 src="images/face/pic6.gif" width=60>
                  <input type="radio" value="7" name="pic">
                  <img border=0 src="images/face/pic7.gif" width=60>
                  <input type="radio" value="8" name="pic">
                  <img border=0 src="images/face/pic8.gif" width=60>
                  <input type="radio" value="9" name="pic">
                  <img border=0 src="images/face/pic9.gif" width=60>
                  <input type="radio" value="10" name="pic">
                  <img border=0 src="images/face/pic10.gif" width=60> </td>
              </tr>
              <tr bgColor="#CCCCCC">
                <td colSpan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                  <input type="submit" value="提交留言" name="Submit">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="reset" value="重新填写" name="Submit2">
                    <input type=hidden name=send value=ok></td>
              </tr>
            </table>
          </form>
    </table></td>
  </tr>
</table>
