<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->

<%
'�˳�����������������ҳ
if request("action")="loginout" then
session.abandon
Response.Redirect ("index.asp")
end if

'������½
if request("login")="check" then
if request("admin")=admin and request("password")=password then
session("login")="ok"
	response.write "<script language='javascript'>"
	response.write "alert('��ӭ���룡������������ǵõ�����˳���������');"
	response.write "location.href='admin.asp';"			
	response.write "</script>"
else
	response.write "<script language='javascript'>"
	response.write "alert('����Ա�û����ƻ�������������������������룡');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
end if
end if

function lleft(content,lef)
for le=1 to len(content)
if asc(mid(content,le,1))<0 then
lef=lef-2
else
lef=lef-1
end if
if lef<=0 then exit for
next
lleft=left(content,le)
end function
%>

<HTML><HEAD>
<TITLE><%=sitename%></TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="description" content="<%=sitename%>">
<meta name="keywords" content="<%=sitename%>">
<link rel="stylesheet" href="book.css" type="text/css">
</HEAD>
<center>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
</head>

<table width="760" border=0 cellspacing=0 cellpadding=0 align=center bgcolor="#FFFFFF" class="grayline">
<tr><td align=center height=50>
  <table border="0" width="100%" id="table2">
	<tr>
		<td background="IMAGES/indextop.gif">
		<p align="center"><a href=admin.asp>������ҳ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; <a href=admin.asp?action=setup>����ѡ��</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=admin.asp?action=loginout>�˳�����</a></td>
	</tr>
</table>
</td></tr>
<tr><td align=center>
<%
if session("login")<>"ok" then 
%>
<br>
<form action="" method=post name=booklogin>
<table width="300" border="1" cellpadding="3" bordercolor="#333333" style="border-collapse: collapse;" align=center>
<tr><td align=center>�û���</td><td><input type="text" name="admin" value="" title="����д�û���"></td></tr>
<tr><td align=center>��&nbsp;&nbsp;��</td><td><input type="password" name="password" value="" title="����д��½����"></td></tr>
<tr><td colspan=2 align=center><input type="submit" value="������½"><input type="hidden" name="login" value="check"></td></tr>
</table>
</form>

<%
else
action=request("action")

'������ҳ
if action="" then%>
<form name=book action=admin.asp method=post>
<table width="600" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#808080" bordercolor="#FFFFFF" bordercolordark="#FFFFFF" style="table-layout:fixed;word-break:break-all">
	<tr><td align=center width=5% height=25>ѡ</td>
	<td align=center width=10%>����</td> 
	<td align=center width=35%>���ݣ��༭��ظ���</td>
	<td align=center width=30%>����</td>
	<td align=center width=11%>״̬</td>
	<td align=center width=9%>���</td>
	</tr>
<%
dim rs,msg_per_page
dim sql
msg_per_page = 10 'ÿҳ��ʾ��¼��
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from Feedback where del=false order by top desc, PostDate desc"

rs.cursorlocation = 3 
rs.pagesize = msg_per_page 'ÿҳ��ʾ��¼��
rs.open sql,conn,1,1 

	if rs.eof and rs.bof then
	response.write "<tr><td colspan=6 align=center height=50>��ʱû������</td></tr>"
	end if

	if not (rs.eof and rs.bof) then '����¼���Ƿ�Ϊ��
		totalrec = RS.RecordCount '�ܼ�¼����
		if rs.recordcount mod msg_per_page = 0 then '������ҳ��,recordcount:���ݵ��ܼ�¼��
		n = rs.recordcount\msg_per_page 'n:��ҳ��
		else 
		n = rs.recordcount\msg_per_page+1 
		end if 

		currentpage = request("page") 'currentpage:��ǰҳ
		If currentpage <> "" then
			currentpage = cint(currentpage)
			if currentpage < 1 then 
				currentpage = 1
			end if 
			if err.number <> 0 then 
				err.clear
				currentpage = 1
			end if
		else
			currentpage = 1
		End if 
		if currentpage*msg_per_page > totalrec and not((currentpage-1)*msg_per_page < totalrec)then 
			currentPage=1
		end if
		rs.absolutepage = currentpage 'absolutepage������ָ��ָ��ĳҳ��ͷ
		rowcount = rs.pagesize 'pagesize������ÿһҳ�����ݼ�¼��
		dim i
		dim k

		Do while not rs.eof and rowcount>0
	content=rs("Comments")	
	replay=rs("replay")
	UserName=rs("UserName")

	if rs("top")="0" then
	Response.write "<tr><td align=center><input type='checkbox' value='"&rs("ID")&"' name=id>"
		else
	Response.write "<tr><td><font color=red title='�̶���Ϣ����ֱ��ɾ�������Ƚ���̶�������ɾ��'>TOP</font>"
	end if
	Response.write "</td><td>"&UserName&"</td><td><a href='admin.asp?action=replay&id="&rs("ID")&"'>"
	response.write lleft(server.htmlencode(content),50)
	response.write "</a></td><td  align=center>"&rs("Postdate")&"</td><td  align=center>"
	if Isnull(Replay) then
		response.write "<font color=red>������</font>"
	else				
		response.write "�ѻظ�"
	end if
		response.write "</td><td align=center>"
		if rs("Online")="0" then response.write "<font color=red>����</font>" else response.write "����"  end if
		response.write "</td></tr>"
	rowcount=rowcount-1
	rs.movenext		
	loop
	end if

	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
%>
<tr><td colspan=6><input type='checkbox' name=chkall onclick='CheckAll(this.form)'> ȫѡ 
	<input type="submit" name="action" value="ɾ��" onClick="{if(confirm('�ò������ɻָ���\n\nȷʵɾ��ѡ�������ԣ�')){this.document.Prodlist.submit();return true;}return false;}"> 	
	</td></tr>
</table>
</form>
<%
call listPages()
end if

if request("action")="ɾ��" then
	delid=replace(request("id"),"'","")
	call delfeedback()
end if

if request("action")="replay" then
	id=request("id")
	call detailfeedback()
end if

if request("action")="setup" then
	call setup()
end if

end if


%>                     
</td></tr>

</table>   

<!--#include file="down.asp"-->
</body></html>

<%
sub setup()
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from book_setup"
rs.open sql,conn,1,3

if request("save")="ok" then

	if trim(request.form("sitename"))="" or trim(request.form("admin"))="" or trim(request.form("maxlength"))="" or trim(request.form("pages"))="" or trim(request.form("huifucolor"))="" or trim(request.form("huifutishi"))="" or trim(request.form("book_jianju"))="" then
	response.write "<script language='javascript'>"
	response.write "alert('������д���������д�������������ύ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if (not isNumeric(request.form("maxlength"))) or (not isNumeric(request.form("pages"))) or (not isNumeric(request.form("book_jianju"))) then
	response.write "<script language='javascript'>"
	response.write "alert('������д���������д�������������ύ��');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if


rs("sitename")=request("sitename")		'վ������
rs("admin")=request("admin")			'����Ա����
if trim(request("password"))<>"" then rs("password")=md5(trim(request("password")))	'����
rs("view")=request("view")			'�������
rs("maxlength")=request("maxlength")		'������󳤶�
rs("pages")=request("pages")			'ÿҳ������
rs("html")=request("html")			'�Ƿ�����html
rs("mailyes")=request("mailyes")		'�Ƿ��������
rs("huifutishi")=trim(request("huifutishi"))	'�ظ���ʾ
rs("huifucolor")=trim(request("huifucolor"))	'�ظ���ɫ
rs("book_jianju")=request("book_jianju")	'������
if request("bad")<>"" then rs("bad")=request("bad")		'�໰����
rs.update

	response.write "<script language='javascript'>"
	response.write "alert('���ύ�������ѱ��档');"
	response.write "location.href='admin.asp?action=setup';"
	response.write "</script>"
	response.end
else
%>
<table width="600" border="1" cellpadding="3" bordercolor="#333333" style="border-collapse: collapse;">
	<form name=book action=admin.asp?action=setup method=post>
	<tr><td  width=25% align=right>��վ���� &nbsp; &nbsp; </td><td><input type=text name=sitename value="<%=rs("sitename")%>" maxlength=20></td></tr>
	<tr><td  width=25% align=right>����Ա���� &nbsp; &nbsp; </td><td><input type=text name=admin value="<%=rs("admin")%>" maxlength=16></td></tr>
	<tr><td  width=25% align=right>����Ա���� &nbsp; &nbsp; </td><td><input type=password name=password value="" maxlength=16> [�����޸����룬������]</td></tr>
	<tr><td  width=25% align=right>���Ա���Email &nbsp; &nbsp; </td><td>��<input type=radio name=mailyes value="0" <%if rs("mailyes")="0" then%> checked<%end if%>> ��<input type=radio name=mailyes value="1" <%if rs("mailyes")="1" then%> checked<%end if%> > [����ʱ�Ƿ������д��ЧE-mail��ַ]</td></tr>
	<tr><td  width=25% align=right>����html���� &nbsp; &nbsp; </td><td>��<input type=radio name=html value="0" <%if rs("html")="0" then%> checked<%end if%>> ��<input type=radio name=html value="1" <%if rs("html")="1" then%> checked<%end if%> > [�Ƿ������û������е�html����]</td></tr>
	<tr><td  width=25% align=right>������˿��� &nbsp; &nbsp; </td><td>��<input type=radio name=view value="0" <%if rs("view")="0" then%> checked<%end if%>> ��<input type=radio name=view value="1" <%if rs("view")="1" then%> checked<%end if%> > [�����Ƿ���Ҫ���]</td></tr>
	<tr><td  width=25% align=right>�������ݳ��� &nbsp; &nbsp; </td><td><input type=text name=maxlength value="<%=rs("maxlength")%>" maxlength=3> [�������ݵ�����ַ���������200]</td></tr>
	<tr><td  width=25% align=right>ÿҳ�������� &nbsp; &nbsp; </td><td><input type=text name=pages value="<%=rs("pages")%>" maxlength=2>  [�Ƽ�10]</td></tr>
	<tr><td  width=25% align=right>���Ա����� &nbsp; &nbsp; </td><td><input type=text name=book_jianju value="<%=rs("book_jianju")%>" maxlength=2> [���ʵļ���ʹ��������ۣ��Ƽ�10]</td></tr>
	<tr><td  width=25% align=right>����Ա�ظ���ʾ &nbsp; &nbsp; </td><td><input type=text name=huifutishi value="<%=rs("huifutishi")%>" maxlength=10> [���磺���ͷ��ظ���]</td></tr>
	<tr><td  width=25% align=right>����Ա�ظ���ɫ &nbsp; &nbsp; </td><td><input type=text name=huifucolor value="<%=rs("huifucolor")%>" maxlength=10> [��ʹ��Ӣ�ĵ��ʻ�16���������磺red��#CCFF00��]</td></tr>
	<tr><td  width=25% align=right>�໰���� &nbsp; &nbsp; </td><td><input type=text name=bad value="<%=rs("bad")%>" size=50 maxlength=255><br> [��/����]</td></tr>
	<tr><td colspan=2><INPUT name="save" TYPE="hidden" value="ok"><INPUT TYPE="submit" value="��������"></td></tr>
	</form>
</table>
<%
end if
rs.close
set rs=nothing
end sub



sub delfeedback()
	if delid="" or isnull(delid) then

	response.write "<script language='javascript'>"
	response.write "alert('����ʧ�ܣ�û��ѡ����ʲ������뵥����ȷ�������أ�');"
	response.write "location.href='admin.asp';"
	response.write "</script>"
	response.end

	else
		conn.execute("delete from Feedback where ID in ("&delid&")")
		conn.close
		set conn=nothing

	response.write "<script language='javascript'>"
	response.write "alert('����ɾ���ɹ����뵥����ȷ�������أ�');"
	response.write "location.href='admin.asp';"
	response.write "</script>"
	response.end

	end if
end sub

sub detailfeedback()
if id="" then
	response.write "<script language='javascript'>"
	response.write "alert('�޴����Ա�ţ��뵥����ȷ�������أ�');"
	response.write "location.href='admin.asp';"
	response.write "</script>"
	response.end
end if

	'�޸���������
if request("send")="ok" then

		if trim(request.form("comments"))="" then 
		response.write "<script language='javascript'>"
		response.write "alert('�����ˣ��������ݲ���Ϊ�գ�');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
		end if

	set rs=server.createobject("adodb.recordset")
	sql = " select * from feedback where del=false and ID="&id
	rs.open sql,conn,1,3

		if not (rs.eof and rs.bof) then
		rs("comments")=request.form("comments")
		rs("Replay")=replace(request.form("Replay"),vbCRLF,"<BR>")
		rs("ReplayDate") = Now()
		rs("Online")=request("Online")
		rs("top")=request("top")
		rs.update
		end if

	rs.close

	response.write "<script language='javascript'>"
	response.write "alert('�����Ѿ��޸Ļ�ظ��ɹ����뵥����ȷ�������أ�');"
	response.write "location.href='admin.asp';"
	response.write "</script>"
	response.end
end if

	'��ʾ��ϸ����
	set rs = server.createobject("adodb.recordset")
	sql = "select * from feedback where ID="&id
	rs.open sql,conn,1,1

		if rs.eof and rs.bof then 
		response.write "<script language='javascript'>"
		response.write "alert('�޴����ԣ��뵥����ȷ�������أ�');"
		response.write "location.href='admin.asp';"
		response.write "</script>"
		response.end
		end if

		if not (rs.eof and rs.bof) then 
			Comments=replace(rs("Comments"),"<BR>",vbCRLF)
			if rs("replay")<>"" then replay=replace(rs("Replay"),"<BR>",vbCRLF) else repley=""  end if

		%>

   <table width="600" border="1" cellpadding="3" bordercolor="#333333" style="border-collapse: collapse;">
		 <form name="repl" method="post" action='admin.asp?action=replay&id=<%=id%>'>
		 <tr><TD align="right" width=20% height=15>������IP��ַ</TD><td><%=rs("IP")%></td></tr>
		 <tr><TD align="right" width=20%>��������</TD><td><%=rs("PostDate")%></td></tr>		 
		 <tr><TD align="right" width=20%>����������</TD><td><%=rs("UserName")%>&nbsp;</td></tr>
		<tr><TD align="right" width=20%>��������</TD><td><%=rs("UserMail")%>&nbsp;</td></tr>
		<tr><TD align="right" width=20%>������ַ</TD><td><%=rs("url")%>&nbsp;</td></tr>
		<tr><TD align="right" width=20%>������ϵ��ʽ</TD><td><%=rs("qq")%>&nbsp;</td></tr>
		 <tr><TD align="right" width=20%>����</TD><td><textarea style="overflow:auto" name="comments" cols="60" rows="8"><%=Comments%></textarea></td></tr>
		 <tr><TD align="right" width=20% valign=top>�ظ�����</TD><td><textarea style="overflow:auto" name="Replay" cols="60" rows="8"><%=replay%></textarea>&nbsp;</td></tr>

		<tr><TD align="right" width=20%>�Ƿ�̶�</TD><td><input type="radio" name="top" value="1" <%if rs("top")="1" then%>checked<%end if%>>
			  �̶�<input type="radio" name="top" value="0" <%if rs("top")="0" then%>checked<%end if%>>
			  ��ͨ </td></tr>


		<tr><TD align="right" width=20%>�Ƿ�����</TD><td><input type="radio" name="Online" value="0" <%if rs("Online")="0" then%>checked<%end if%>>
			  ����<input type="radio" name="Online" value="1" <%if rs("Online")="1" then%>checked<%end if%>>
			  ���� </td></tr>
			<TR><TD align="right" width=20%>&nbsp;<INPUT TYPE="hidden" name=send value=ok></TD><TD>
				<input type="submit" name="action" value=" �� �� "></TD></TR>
			</form></TABLE>
		<%
		end if	

	rs.close
	set rs=nothing

end sub


'��ҳ
sub listPages() 
if n <= 1 then exit sub 
%>
��<%=totalrec%>������ 
<%if currentpage = 1 then%>
<font color=darkgray>��ҳ ǰҳ</font>
<%else%> 
<a href="<%=request.ServerVariables("script_name")%>?page=1">
��ҳ</font></a> <a href="<%=request.ServerVariables("script_name")%>?page=<%=currentpage-1%>">ǰҳ</a>
<%end if%>
<%if currentpage = n then%> 
<font color=darkgray >��ҳ ĩҳ</font>
<%else%> 
<a href="<%=request.ServerVariables("script_name")%>?page=<%=currentpage+1%>">��ҳ</a> <a href="<%=request.ServerVariables("script_name")%>?page=<%=n%>">ĩҳ</a>
<%end if%>
  ��<%=currentpage%>ҳ ��<%=n%>ҳ
<%end sub%>