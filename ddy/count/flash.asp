<%@ CODEPAGE = "936" %>
<!--#include file="inc_config.asp"-->
<%
on error resume next
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from vjian"
rs.Open sql,conn,1,1

vuser=cint(Request.Cookies(mNameEn)("lao"))+1

Response.Write "total=" & rs("vtop") & "&yes=" & rs("yesterday") & "&today=" & _
	rs("today") & "&you=" & vuser & "&load=end"

rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>