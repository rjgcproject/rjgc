<%
style=Request("style")
theurl="http://" & Request.ServerVariables("http_host") & finddir(Request.ServerVariables("url"))
%>
document.write("<script>var style='<%=style%>';var url='<%=theurl%>';</script><script src='<%=theurl%>stat.js'></script>")
<%
Function finddir(filepath)
	finddir=""
	for i=1 to len(filepath)
	if left(right(filepath,i),1)="/" or left(right(filepath,i),1)="\" then
	  abc=i
	  exit for
	end if
	next
	if abc <> 1 then
	finddir=left(filepath,len(filepath)-abc+1)
	end if
end Function
%>
