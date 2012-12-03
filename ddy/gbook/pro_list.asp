<table width="193" border="0" align="center" cellpadding="0" cellspacing="0" class="zi">
  <tr>
    <td><img src="../images/products_list.jpg" width="193" height="55" /></td>
  </tr>
  <tr>
    <td><table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><table border="0" cellpadding="0" cellspacing="0" 
            width="100%">
            <!--DWLayoutTable-->
            <tbody>
              <tr>
                <td valign="top"><%
        Set rs=Server.CreateObject("ADODB.RecordSet")
        strsql="select * from class"
		rs.Open strsql,con,1,1
		if rs.eof then 
    	else
        rs.AbsolutePage=pagecount
    	  %>
                    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <%
          j=1
          do while not (rs.eof or err)
          if (j mod 1)=1 then
              response.write "<tr>"
          end if
        %>
                      <tr>
                        <td align="center" valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr>
                              <td align="center" valign="middle"><table width="100%"  border="0" align="left" cellpadding="0" cellspacing="0" >
                                  <tr onMouseOver="this.bgColor='#F0F0F0'" onMouseOut="this.bgColor=''">
                                    <td width="5%" align="center">¡¤</td>
                                    <td width="4%" align="center">&nbsp;</td>
                                    <td width="91%" height="30" align="center"><div align="left"><a href="../products.asp?Sort_id=<%=rs("Sort_id")%>"><%=left(rs("Sort"),12)%></a></div></td>
                                  </tr>
                              </table></td>
                            </tr>
                        </table></td>
                        <%  
          if (j mod 7)=0 then
		     response.write "</tr>"
          end if
	      j=j+1
		  rs.moveNext
		  if j>16 then exit do
          loop
    end if
       %>
                      </tr>
                    </table></td>
              </tr>
            </tbody>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="193" border="0" align="center" cellpadding="0" cellspacing="0" class="zi">
      <tr>
        <td width="100%"><a href="../contact.asp"><img src="../images/left_ser.gif" alt="&cedil;&uuml;&para;&agrave;&Aacute;&ordf;&Iuml;&micro;&middot;&frac12;&Ecirc;&frac12;" width="192" height="100" border="0" /></a></td>
      </tr>
      <tr>
        <td><a href="../joinus.asp"><img src="../images/left_join.jpg" alt="&frac14;&Oacute;&Egrave;&euml;&Icirc;&Ograve;&Atilde;&Ccedil;" width="192" height="100" border="0" /></a></td>
      </tr>
    </table></td>
  </tr>
  
</table>
