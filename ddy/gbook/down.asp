<%

set rs=conn.execute("select count(*) as offnum from Feedback where online<>'1'")
if cint(rs("offnum"))>0 then tit="£¬¹²ÓÐ "&rs("offnum")&" ÌõÎ´ÉóºËÁôÑÔ"
set rsoff=nothing

%>
<style type="text/css">
<!--
.style6 {color: #999999}
.style7 {font-family: Arial, Helvetica, sans-serif}
-->
</style>
<table width="1003" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="1" bgcolor="#003366"></td>
  </tr>
  <tr>
    <td height="50" class="style1">
      <div align="right"><span class="style6">&nbsp;<span class="style7"> </span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></div></td>
  </tr>
</table>
