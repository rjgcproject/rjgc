<%
option explicit
response.buffer=true
%>
<!--#include file="inc/setup.asp"-->
<!--#include file="inc/cj_cls.asp"-->
<%
'-------------------------------------
Dim Rs,Sql,action
dim ChannelID,channelname,channelshortname,id,i,ParentID,depth,classID,Colleclx,href
Colleclx=request("Colleclx")
if Colleclx="" then Colleclx=0
Dim CjName : CjName=Skcj.GetItemConfig("CjName",Colleclx)
Call Skcj.Show_Top()
%>
<html>
<head>
<title>�ɼ�ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/Admin_Style.css">
<style type="text/css">
.ButtonList {
	BORDER-RIGHT: #000000 2px solid; BORDER-TOP: #ffffff 2px solid; BORDER-LEFT: #ffffff 2px solid; CURSOR: default; BORDER-BOTTOM: #999999 2px solid; BACKGROUND-COLOR: #e6e6e6
}
</style>
<SCRIPT language=javascript>
    function unselectall(thisform){
        if(thisform.chkAll.checked){
            thisform.chkAll.checked = thisform.chkAll.checked&0;
        }   
    }
    function CheckAll(thisform){
        for (var i=0;i<thisform.elements.length-6;i++){
            var e = thisform.elements[i];
            if (e.Name !="chkAll"&&e.disabled!=true)
                e.checked = thisform.chkAll.checked;
        }
    }
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>����������</strong></td>
    <td height="30"><a href="sk_class.asp?action=add&Colleclx=<%=Colleclx%>">��Ok3w����ϵͳ�е������</a></td>
  </tr>
</table>
<% 
action =request("action")
select case action
case "add"
	if request("add_ok")<>"" and request("add_ok")<>0 then
		Call SK_Class_Channel_add()
	else
		call add()
	end if
case else
	call main()
end select

sub main()
 %>
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder" >
    <tr> 
      <th height="22" colspan="2" class="title"> <div align="center"><b>����&nbsp;��&nbsp;��</b></div></th>
    </tr>
    <tr>
      <td height="22" colspan="2" ><table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
        <TR>
          <TD width="50" bgcolor="#CAD7F7"> ���</TD>
          <TD bgcolor="#CAD7F7"> ��Ŀ��������</TD>
        </TR>
        <%Call DisClassTable(1,0)%>
      </TABLE></td>
    </tr>
</table>
<% end sub %>
<% sub add() %>
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" class="border" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><b></b></div></td>
    </tr>
</table>
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
    <TR> 
      <TD height=22  align=center >��OK3W����ϵͳ�е������</TD>
    </TR>
</TABLE>
  <form name="myform" method="POST" action="?action=add&add_ok=1&Colleclx=<%=Colleclx%>">
<div align="center"> <%=ErrMsg%>
  <input type="submit" name="Submit" value="��������µ�����ࣨ���ڷ��ཫ��գ�" onClick="if(!confirm('���Ҫ���µ�����')) return false;">
</div>
  </form>
<% end sub %>

</body>
</html>
<%
Call CloseConnItem()
%>