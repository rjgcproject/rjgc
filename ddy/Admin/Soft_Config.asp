<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%dbdns="../"%>
<!--#include file="chk.asp"-->
<!--#include file="../AppCode/Class/Ok3w_SiteConfig.asp"-->
<!--#include file="../AppCode/fun/function.asp"-->
<%
Call CheckAdminFlag(6)

Select Case Trim(Request.Form("action"))
	Case "edit"
		Set SiteConfig = New Ok3w_SiteConfig
		Call SiteConfig.SoftEdit()
		Call SaveAdminLog("�༭Soft����")
		Call CloseConn()
		Call ActionOk("Soft_Config.asp")
End Select
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̨����ϵͳ</title>
<link rel="stylesheet" type="text/css" href="images/Style.css">
</head>

<body>
<!--#include file="top.asp"-->
<br />
<table cellspacing="0" cellpadding="0" width="98%" align="center" border="0">
  <tbody>
    <tr>
      <td style="PADDING-LEFT: 2px; HEIGHT: 22px" 
    background="images/tab_top_bg.gif"><table cellspacing="0" cellpadding="0" width="477" border="0">
        <tbody>
          <tr>
            <td width="147"><table height="22" cellspacing="0" cellpadding="0" border="0">
              <tbody>
                <tr>
                  <td width="3"><img id="tabImgLeft__0" height="22" 
                  src="images/tab_active_left.gif" width="3" /></td>
                  <td 
                background="images/tab_active_bg.gif" class="tab"><strong class="mtitle">��̨����ϵͳ</strong></td>
                  <td width="3"><img id="tabImgRight__0" height="22" 
                  src="images/tab_active_right.gif" 
            width="3" /></td>
                </tr>
              </tbody>
            </table></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr>
      <td bgcolor="#ffffff"><table cellspacing="0" cellpadding="0" width="100%" border="0">
        <tbody>
          <tr>
            <td width="1" background="images/tab_bg.gif"><img height="1" 
            src="images/tab_bg.gif" width="1" /></td>
            <td 
          style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
          valign="top"><div id="tabContent__0" style="DISPLAY: block; VISIBILITY: visible">
              <table cellspacing="1" cellpadding="1" width="100%" align="center" 
            bgcolor="#8ccebd" border="0">
                <tbody>
                  <tr>
                    <td 
                style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; PADDING-TOP: 10px" 
                valign="top" bgcolor="#fffcf7"><table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                      <form id="form1" name="form1" method="post" action="">
                        
                        <tr>
                          <td colspan="2" align="left" bgcolor="#EBEBEB"><strong>����������������</strong></td>
                        </tr>
                        <tr>
                          <td width="18%" align="right" bgcolor="#FFFFFF">����Ƶ������</td>
                          <td width="82%" bgcolor="#FFFFFF"><input name="SiteSoftTitle" type="text" id="SiteSoftTitle" value="<%=Application("Ok3w_SiteSoftTitle")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">����Ƶ���ؼ���</td>
                          <td bgcolor="#FFFFFF"><input name="SiteSoftKeyWords" type="text" id="SiteSoftKeyWords" value="<%=Application("Ok3w_SiteSoftKeyWords")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">����Ƶ��˵��</td>
                          <td bgcolor="#FFFFFF"><input name="SiteSoftDescription" type="text" id="SiteSoftDescription" value="<%=Application("Ok3w_SiteSoftDescription")%>" size="50" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">���ط�����</td>
                          <td bgcolor="#FFFFFF"><textarea name="SiteSoftServer" cols="80" rows="8" id="SiteSoftServer"><%=Application("Ok3w_SiteSoftServer")%></textarea>
                            <br>
                            ����ж�������ûس��ָ���һ��һ������|���ָ����������ƺͷ�������ַ��ÿһ�б����ԡ�/������</td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">Ѹ����������ID</td>
                          <td bgcolor="#FFFFFF"><input name="SiteSoftXunlei" type="text" id="SiteSoftXunlei" value="<%=Application("Ok3w_SiteSoftXunlei")%>" size="15" />
                            Ϊ�ձ�ʾ������</td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">�쳵��������ID</td>
                          <td bgcolor="#FFFFFF"><input name="SiteSoftFlashget" type="text" id="SiteSoftFlashget" value="<%=Application("Ok3w_SiteSoftFlashget")%>" size="15" />
                            Ϊ�ձ�ʾ������</td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">��������</td>
                          <td bgcolor="#FFFFFF"><input name="SiteSoftLanguage" type="text" id="SiteSoftLanguage" value="<%=Application("Ok3w_SiteSoftLanguage")%>" size="82" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">&nbsp;������Ȩ</td>
                          <td bgcolor="#FFFFFF"><input name="SiteSoftLicence" type="text" id="SiteSoftLicence" value="<%=Application("Ok3w_SiteSoftLicence")%>" size="82" /></td>
                        </tr>
                        

                        
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">��������</td>
                          <td bgcolor="#FFFFFF"><input name="SiteSoftProperty" type="text" id="SiteSoftProperty" value="<%=Application("Ok3w_SiteSoftProperty")%>" size="82" /></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">����ϵͳ</td>
                          <td bgcolor="#FFFFFF"><input name="SiteSoftTos" type="text" id="SiteSoftTos" value="<%=Application("Ok3w_SiteSoftTos")%>" size="82" /></td>
                        </tr>
                        
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">����˵��</td>
                          <td bgcolor="#FFFFFF"><textarea name="SiteSoftDownHelp" cols="80" rows="8" id="SiteSoftDownHelp"><%=Application("Ok3w_SiteSoftDownHelp")%></textarea></td>
                        </tr>
                        <tr>
                          <td align="right" bgcolor="#FFFFFF">&nbsp;</td>
                          <td bgcolor="#FFFFFF"><input name="action" type="hidden" id="action" value="edit" />
                              <input type="submit" name="Submit" value="�� ��" style="font-size:14px;" /></td>
                        </tr>
                      </form>
                    </table></td>
                  </tr>
                </tbody>
              </table>
            </div></td>
            <td width="1" background="images/tab_bg.gif"><img height="1" 
            src="images/tab_bg.gif" width="1" /></td>
          </tr>
        </tbody>
      </table></td>
    </tr>
    <tr>
      <td background="images/tab_bg.gif" bgcolor="#ffffff"><img height="1" 
      src="images/tab_bg.gif" width="1" /></td>
    </tr>
  </tbody>
</table>
</body>
</html>
<%
Call CloseConn()
Set Admin = Nothing
Set SiteConfig = Nothing
%>
