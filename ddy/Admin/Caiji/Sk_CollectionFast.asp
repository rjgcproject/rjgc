<%
option explicit
Response.Buffer = True 
Server.ScriptTimeOut=999
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache"
%>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/SK_FunCls.asp"-->
<!--#include file="inc/clsCache.asp"-->
<!--#include file="inc/cj_cls.asp"-->
<%
dim Skcj
Set Skcj= New FunCls
Dim Action,CollecType
Dim myCache
Dim ItemNum,ListNum,PaingNum,NewsSuccesNum,NewsFalseNum,NewsNum_i,Itemon,ItemIdstr,Itemok
Dim Rs,Sql,RsItem,SqlItem,FoundErr,ItemEnd,ListEnd
Dim PicUrls_i,NewsUrlPaing_s,NewsUrlPaing_o,NewsPaingNext_Code,TypeArray_Url,TypeNews_Url

'��Ŀ����
Dim ItemID,ItemName,ChannelID,strChannelDir,ClassID,SpecialID,LoginType,LoginUrl,LoginPostUrl,LoginUser,LoginPass,LoginFalse
Dim ListStr,LsString,LoString,ListPaingType,LPsString,LPoString,ListPaingStr1,ListPaingStr2,ListPaingID1,ListPaingID2,ListPaingStr3,HsString,HoString,HttpUrlType,HttpUrlStr,photourls,photourlo,PhotoPaingType,PhotoType_s,PhotoType_o,PhotoLurl_s,PhotoLurl_o,Phototypefy_s,Phototypefy_o,Phototypefyurl_s,Phototypefyurl_o,Phototypeurl_s,Phototypeurl_o,Colleclx,selEncoding,SaveFileUrl,x_tpUrl,Thumb_WaterMark,Thumbs_Create,Timing,strReplace

'���ر���
dim DownSize,DownYY,DownSQ,DownPT,YSDZ,ZCDZ,PhotoUrl,DownUrls
'���ر�����Ŀ�ֶ�
dim Downlist_s,Downlist_o,DownUrl_s,DownUrl_o,DownNewType,DownNewlist_s,DownNewlist_o,DownNewUrl_s,DownNewUrl_o,LinkUrlYn
dim ZdType_001,Zds_001,Zdo_001,ZD_001,ZdType_002,Zds_002,Zdo_002,ZD_002,ZdType_003,Zds_003,Zdo_003,ZD_003,ZdType_004,Zds_004,Zdo_004,ZD_004,ZdType_005,Zds_005,Zdo_005,ZD_005,ZdType_006,Zds_006,Zdo_006,ZD_006,ZdType_007,Zds_007,Zdo_007,ZD_007,ZdType_008,Zds_008,Zdo_008,ZD_008

'--ͼƬ�б�����
dim imhstr,imostr,NewsimageCode,Newsimage,picpath,Radiobutton,x_tp
'--ͼƬ�б�����
Dim TsString,ToString,CsString,CoString,DateType,DsString,DoString,AuthorType,AsString,AoString,AuthorStr,CopyFromType,FsString,FoString
Dim CopyFromStr,KeyType,KsString,KoString,KeyStr,NewsPaingType,NPsString,NpoString,NewsPaingStr,NewsPaingHtml
Dim ItemCollecDate,PaginationType,MaxCharPerPage,ReadLevel,Stars,ReadPoint,Hits,UpDateType,UpDateTime,IncludePicYn,DefaultPicYn,OnTop,Elite,Hot
Dim SkinID,TemplateID,Script_Iframe,Script_Object,Script_Script,Script_Div,Script_Class,Script_Span,Script_Img,Script_Font,Script_A,Script_Html,CollecListNum,CollecNewsNum,Passed,SaveFiles,CollecOrder,InputerType,Inputer,EditorType,Editor,ShowCommentLink,Script_Table,Script_Tr,Script_Td

'���˱���
Dim Arr_Filters,FilterStr,Filteri

'�ɼ���صı���
Dim ContentTemp,NewsPaingNext,NewsPaingNextCode,Arr_i,NewsUrl,NewsCode,ListTypeCode,ListTypeUrlCode,TypeUrlArray,TypeNewsUrl,NewsTypeCode,PicUrls,Arr_ii,Arr_ii_2,ListTypeCode_2,ListTypeUrlCode_2,TypeUrlArray_2

'���±������
Dim ArticleID,Title,Content,Author,CopyFrom,Key,IncludePic,UploadFiles,DefaultPicUrl,Coll_DefiniteUrl
'��������
Dim LoginData,LoginResult,OrderTemp,i
Dim Arr_Item,CollecTest,Content_View,CollecNewsAll
Dim StepID

'��ʷ��¼
Dim Arr_Histrolys,His_Title,His_CollecDate,His_Result,His_Repeat,His_i 

'ִ��ʱ�����
Dim StartTime,OverTime

'ͼƬͳ��
Dim Arr_Images,ImagesNum,ImagesNumAll

'�б�
Dim ListUrl,ListCode,NewsArrayCode,NewsArray,ListArray,ListPaingNext

'��װ·��
Dim strInstallDir,CacheTemp
Dim DiyFieldSTR_z,DiyFieldSTR_l'�Զ���
Dim FoundErr_1
'----�Ƿ��½---------------------
Call Skcj.Admin()
If Skcj.IsAdmin=False Then
	ErrMsg="<li> ��û�е�½���ǹ���Ա����<a href='sk_login.asp' target='_top'>��½</a>��"
	response.Redirect("Sk_err.asp?action=AdminErr&ErrMsg="&ErrMsg&"")
	response.End()
End If
'-------------------------------------
'On Error Resume Next
strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
'����·��
CacheTemp=Lcase(trim(request.ServerVariables("SCRIPT_NAME")))
CacheTemp=left(CacheTemp,instrrev(CacheTemp,"/"))
CacheTemp=replace(CacheTemp,"\","_")
CacheTemp=replace(CacheTemp,"/","_")
CacheTemp="ansir" & CacheTemp

'���ݳ�ʼ��
CollecListNum=0'�б���
CollecNewsNum=0'�б���
ArticleID=0'ID
ItemNum=Clng(Trim(Request("ItemNum")))
ListNum=Clng(Trim(Request("ListNum")))
NewsNum_i=Clng(Trim(Request("NewsNum_i")))
NewsSuccesNum=Clng(Trim(Request("NewsSuccesNum")))
NewsFalseNum=Clng(Trim(Request("NewsFalseNum")))
ImagesNumAll=Clng(Trim(Request("ImagesNumAll")))
ListPaingNext=Trim(Request("ListPaingNext"))
Itemon=Trim(Request("Itemon"))'���ٲɼ�
Itemok=Trim(Request("Itemok"))'���ٲɼ�
FoundErr=False
ItemEnd=False
ListEnd=False
ErrMsg=""

Call DelNews()'
Call CheckForm()'���ItemIDֵ
Dim Collecdate : Collecdate=Trim(Request("Collecdate"))
If Itemok = "yes" then 
	If Instr(Itemon,",")>0 Then
		ItemIdstr=GetItemId(Itemon,1)
		Response.write("<script>location.href='Sk_CollectionFast.asp?ItemID="&GetItemId(Itemon,0)&"&ItemNum=1&ListNum=1&NewsSuccesNum=0&NewsFalseNum=0&ImagesNumAll=0&NewsNum_i=0&Itemon="& ItemIdstr &"&Collecdate="& Collecdate &"';</script>")'��ҳ��
	Else
		Response.write("<script>location.href='Sk_CollectionFast.asp?ItemID="&Itemon&"&ItemNum=1&ListNum=1&NewsSuccesNum=0&NewsFalseNum=0&ImagesNumAll=0&NewsNum_i=0&Collecdate="& Collecdate &"';</script>")'��ҳ��
	End if
	Response.end	
End if

If Instr(ItemID,",")>0 Then
	ItemIdstr=GetItemId(ItemID,1)
	Response.write("<script>location.href='Sk_CollectionFast.asp?ItemID="&GetItemId(ItemID,0)&"&ItemNum=1&ListNum=1&NewsSuccesNum=0&NewsFalseNum=0&ImagesNumAll=0&NewsNum_i=0&Itemon="& ItemIdstr&"&Collecdate="& Collecdate &"';</script>")'��ҳ��
	Response.end
End if

If FoundErr<>True Then
  Call SetCache()'��Ŀ��Ϣд�뻺��
End If
If FoundErr=True Then
   Call WriteErrMsg(ErrMsg)
Else
   Call GetCache()
   Call Main()
   Collection_Fast
End If
'�ر����ݿ�����
Call CloseConnItem()
%>
<%Sub Main%>
<html>
<head>
<title>�ɼ�ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/Admin_Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<br>      
</body>         
</html>
<%End Sub

Sub CheckForm()'��ȡ����
   ItemID=Trim(Request("ItemID"))
   CollecTest=Trim(Request.Form("CollecTest"))
   Content_View=Trim(Request.Form("Content_View"))
   '������
   If ItemID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>����ѡ����Ŀ!</li>"
   Else
      If Instr(ItemID,",")>0 Then
         ItemID=Replace(ItemID," ","")
      End If
	  Response.Flush()
	  set rs=connItem.execute("select top 1 * from Item Where ItemID in(" & ItemID &")" )
   	  if connItem.Execute("select count(ClassID) from SK_Class Where ClassID=" & RS("ClassID"))(0)=0 then
	  'if conn.Execute("select count(id) from ks_Class Where ID='" & RS("ClassID") &"'")(0)=0 then
	  FoundErr=True
	  ErrMsg=ErrMsg & "<br><li>��������Ƶ����Ŀ! </li>"
	  rs.close
	  set rs=nothing
	  end if
   End If 
   If CollecTest="yes" Then
      CollecTest=True
   Else
      CollecTest=False
   End If
   If Content_View="yes" Then
      Content_View=True
   Else
      Content_View=False
   End If
End Sub

'==================================================
'��������SetCache1
'��  �ã���ȡ����
'��  ������
'==================================================
Sub GetCache()
   Dim myCache
   Set myCache=new SK_clsCache

   '��Ŀ��Ϣ
   myCache.name=CacheTemp & "items"
   If myCache.valid then 
      Arr_Item=myCache.value
   Else
      ItemEnd=True
   End If

   '������Ϣ
   myCache.name=CacheTemp & "filters"
   If myCache.valid then 
      Arr_Filters=myCache.value
   End If

   '��ʷ��¼
   myCache.name=CacheTemp & "histrolys"
   If myCache.valid then 
      Arr_Histrolys=myCache.value
   End If

   '������Ϣ
   myCache.name=CacheTemp & "collectest"
   If myCache.valid then 
      CollecTest=myCache.value
   Else
      CollecTest=False
   End If
   myCache.name=CacheTemp & "contentview"
   If myCache.valid then 
      Content_View=myCache.value
   Else
      Content_View=False
   End If

   Set myCache=Nothing
End Sub

Sub SetCache()'��Ŀ��Ϣд�뻺��
   SqlItem ="select * from Item where ItemID in(" & ItemID & ")"
   Set RsItem=Server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If Not RsItem.Eof Then
      Arr_Item=RsItem.GetRows()
   End If
   RsItem.Close
   Set RsItem=Nothing

   Set myCache=new SK_clsCache
   myCache.name=CacheTemp & "items"
   Call myCache.clean()
   If IsArray(Arr_Item)=True Then
      myCache.add Arr_Item,Dateadd("n",1000,now)
   Else
      FoundErr=True
      ErrMsg=ErrMsg & "<br>�����������"
   End If

   '������Ϣ
   SqlItem ="select * from Filters where Flag=True"
   Set RsItem=Server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If Not RsItem.Eof Then
      Arr_Filters=RsItem.GetRows()
   End If
   RsItem.Close
   Set Rsitem=Nothing

   myCache.name=CacheTemp & "filters"
   Call myCache.clean()
   If IsArray(Arr_Filters)=True Then
      myCache.add Arr_Filters,Dateadd("n",1000,now)
   End If

   '��ʷ��¼
   SqlItem ="select NewsUrl,Title,CollecDate,Result from Histroly"
   Set RsItem=Server.CreateObject("adodb.recordset")
   RsItem.Open SqlItem,ConnItem,1,1
   If Not RsItem.Eof Then
      Arr_Histrolys=RsItem.GetRows()
   End If
   RsItem.Close
   Set RsItem=Nothing

   myCache.name=CacheTemp & "histrolys"
   Call myCache.clean()
   If IsArray(Arr_Histrolys)=True Then
      myCache.add Arr_Histrolys,Dateadd("n",1000,now)
   End If

   '������Ϣ
   myCache.name=CacheTemp & "collectest"
   Call myCache.clean()
   myCache.add CollecTest,Dateadd("n",1000,now)

   myCache.name=CacheTemp & "contentview"
   Call myCache.clean()
   myCache.add Content_View,Dateadd("n",1000,now)

   set myCache=nothing
End Sub
Sub DelNews()
   ConnItem.execute("Delete From [NewsList]")
End Sub

'=============���ٲɼ�==================	
Public Sub Collection_Fast
	If ItemEnd<>True Then
	   If (ItemNum-1)>Ubound(Arr_Item,2) then
		  ItemEnd=True
	   Else
		  SK.SetItems()
	   End If
	End If
	If ItemEnd<>True Then
	   If ListPaingType=0 Then
		  If ListNum=1 Then
			ListUrl=ListStr
		  Else
			 ListEnd=True
		  End If
	   ElseIf ListPaingType=1 Then
		if listnum=1 and ListPaingID1<>1 and  ListPaingID2<>1 then 
			ListUrl=ListStr
		   else
			  If ListPaingID1>ListPaingID2 then
				 If (ListPaingID1-ListNum+1)<ListPaingID2 or (ListPaingID1-ListNum+1)<0 Then
					Listend=True
				 Else
					ListUrl=Replace(ListPaingStr2,"{$ID}",Cstr(ListpaingID1-ListNum+1))
				 End if
			  Else
				 If (ListPaingID1+ListNum-1)>ListPaingID2 Then
					ListEnd=True
				 Else
					ListUrl=Replace(ListPaingStr2,"{$ID}",CStr(ListPaingID1+ListNum-1))
				 End If
			  End If   
		 end if   
	   ElseIf ListPaingType=2 Then
		  ListArray=Split(ListPaingStr3,"|")
		  If (ListNum-1)>Ubound(ListArray) Then
			 ListEnd=True
		  Else
			 ListUrl=ListArray(ListNum-1)
		  End If
	   ElseIf ListPaingType=3 Then
		  If ListNum=1 Then
			ListUrl=ListStr
			ListCode=ListStr
			application.Lock()
			application("IPQ_SS")=ListUrl
			application.UnLock()
		  Else
			ListUrl=application("IPQ_SS")
			ListCode=SKcj.GetHttpPage(ListUrl,"GB2312")
			ListCode=SKcj.GetBody(ListCode,LPsString,LPoString,False,False)
		  End if
		  If ListCode<>"$False$" or ListCode<>"" Then
			 ListCode=Trim(Skcj.FormatRemoteUrl(ListCode,ListStr))
			 ListUrl=ListCode
			 application.Lock()
			 application("IPQ_SS")=ListUrl
			 application.UnLock()
		  Else
			 ListEnd=True
		  End if    	  
	   End If
	   
	   If ListNum>CollecListNum And CollecListNum<>0 Then
		  ListEnd=True
	   End if
	End If
	
	
	If ItemEnd=True Then
	   ErrMsg= "<br>�ɼ�����ȫ�����"
	   ErrMsg=ErrMsg & "<br>�ɹ��ɼ��� "  &  NewsSuccesNum  &  "  ��,ʧ�ܣ� "    &  NewsFalseNum  &  "  ��,ͼƬ��" & ImagesNumAll & "  ��"
	   Call DelCache()
	   ErrMsg=ErrMsg & "<meta http-equiv=""refresh"" content=""1;url="& Skcj.GetItemConfig("FileName",Colleclx) &""">"
	Else
	   If ListEnd=True Then
	      If Itemon<>"" then  Itemok= "yes"'ȫѡ�ɼ�
		  If Instr(Itemon,",")>0 or Itemon<>"" Then
			  if Collecdate<>"" Then 
					Collecdate=Day(now())
					response.write("<script>location.href='sk_Timing.asp?action=GoTiming&Collecdate="&  Day(now()) &"';</script>")'��ҳ��
			  Else
					response.write("<script>location.href='Sk_CollectionFast.asp?ItemID="& ItemID &"&ItemNum=" & ItemNum & "&ListNum=0&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ListPaingNext=" & ListPaingNext &"&NewsNum_i="& 0 &"&Itemok="& Itemok &"&Itemon="& Itemon &"&Collecdate="& Collecdate &"';</script>")'���
					Response.end
			  End if	  
		  End if
		  ItemNum=ItemNum+1
		  ListNum=1
		  ErrMsg="<br>" & ItemName & "  ��Ŀ�����б��ɼ���ɣ����������������Ժ�..."
		  ErrMsg=ErrMsg & "<meta http-equiv=""refresh"" content=""1;url="& Skcj.GetItemConfig("FileName",Colleclx) &""">"
	   End If
	End If
	
	SK.TopItem()'Top����
	  If ItemEnd=True Or ListEnd=True Then
		If ItemEnd<>True Then
		  Sk.SetCache_His()
		End If
		Call WriteSucced(ErrMsg)
	Else
	   FoundErr=False
	   ErrMsg=""
	   Call StartCollection()'��ʼ�ɼ�
	   SK.FootItem2()
	End  If	
End Sub

Sub StartCollection'��ʼ�ɼ�
IF Colleclx <> 0 then 
	Set Rs = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID="& Colleclx )
Else
	Set Rs = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID=1" )
End if
Skcj.CjTimeout=Rs("Timeout") 
Skcj.DownExtName=Rs("FileExtName")
Skcj.MaxSize=Rs("MaxFileSize")
Rs.close : Set Rs=Nothing
If NewsSuccesNum >= CollecNewsNum And CollecNewsNum<>0 then 
	 If Itemon="" then
	  	  if Collecdate<>"" then
			 response.write("<script>location.href='sk_Timing.asp?action=GoTiming&Collecdate="&  Day(now()) &"';</script>")
		  Else
			  Response.Write "<br> &nbsp;&nbsp;&nbsp;&nbsp;�ɼ���ɣ����������������Ժ�..." 
			  Response.Write  "<meta http-equiv=""refresh"" content=""1;url="& Skcj.GetItemConfig("FileName",Colleclx) &""">"
	  	  End if
	  Else
		 response.write  "<script>location.href='Sk_CollectionFast.asp?ItemID="& ItemID &"&ItemNum=1&ListNum=1&NewsSuccesNum=0&NewsFalseNum=0&ImagesNumAll=0&NewsNum_i=0&Itemon="& Itemon &"&Itemok=yes&Collecdate="& Collecdate &"';</script>"
	  End if
	  Response.end
End if
If FoundErr<>True then
   ListCode=Skcj.ReplaceTrim(Skcj.GetHttpPage(ListUrl,selEncoding))
   Sk.GetListPaing()
   If ListCode="$False$" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ�б���" & ListUrl & "��ҳԴ��ʱ��������</li>"
   Else
      ListCode=Skcj.GetBody(ListCode,LsString,LoString,False,False)
      If ListCode="$False$" Or ListCode="" Then
         FoundErr=True
		 FoundErr_1=True
         ErrMsg=ErrMsg & "<br><li>�ڽ�ȡ��" & ListUrl & "��"& Skcj.GetItemConfig("CjName",Colleclx) &"�б�ʱ��������</li>"
      End If
   End If
End If

If FoundErr<>True Then
   NewsArrayCode=Skcj.GetArray(ListCode,HsString,HoString,False,False)
   If NewsArrayCode="$False$" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>�ڷ�����" & ListUrl & ""& Skcj.GetItemConfig("CjName",Colleclx) &"�б�ʱ��������</li>"
   Else
      NewsArray=Split(NewsArrayCode,"$Array$")
      For Arr_i=0 to Ubound(NewsArray)
         If HttpUrlType=1 Then
            NewsArray(Arr_i)=Trim(Replace(HttpUrlStr,"{$ID}",NewsArray(Arr_i)))
         Else
            NewsArray(Arr_i)=Trim(FormatRemoteUrl(NewsArray(Arr_i),ListUrl))    
         End If
         NewsArray(Arr_i)=CheckUrl(NewsArray(Arr_i))
      Next 
      If CollecOrder=1 Then
         For Arr_i=0 to Fix(Ubound(NewsArray)/2)
            OrderTemp=NewsArray(Arr_i)
            NewsArray(Arr_i)=NewsArray(Ubound(NewsArray)-Arr_i)
            NewsArray(Ubound(NewsArray)-Arr_i)=OrderTemp
         Next
      End If
   End If
End If


If FoundErr<>True Then
	If x_tp=1 then
	   NewsimageCode=Skcj.GetArray(ListCode,imhstr,imostr,False,False)
	   If NewsimageCode="$False$" Then
		  FoundErr=True
		  ErrMsg=ErrMsg & "<br><li>�ڷ�����" & ListUrl & "СͼƬ�б�ʱ��������</li>"
	   Else
		  Newsimage=Split(NewsimageCode,"$Array$")
		  For Arr_i=0 to Ubound(Newsimage)
			 If HttpUrlType=1 Then
				Newsimage(Arr_i)=Trim(Replace(HttpUrlStr,"{$ID}",Newsimage(Arr_i)))
			 Else
				Newsimage(Arr_i)=Trim(Skcj.FormatRemoteUrl(Newsimage(Arr_i),ListUrl))           
			 End If
			 if x_tpUrl<>"" then Newsimage(Arr_i)= x_tpUrl & Newsimage(Arr_i) 
			 Newsimage(Arr_i)=CheckUrl(Newsimage(Arr_i))
			 
		  Next
		  If CollecOrder=True Then
			 For Arr_i=0 to Fix(Ubound(Newsimage)/2)
				OrderTemp=Newsimage(Arr_i)
				Newsimage(Arr_i)=Newsimage(Ubound(Newsimage)-Arr_i)
				Newsimage(Ubound(Newsimage)-Arr_i)=OrderTemp
			 Next
		  End If
	   End If
	End If
End if
If FoundErr<>True Then
   dim PicUrls_i
   SK.TopItem2()
   CollecNewsAll=0
   For Arr_i=0 to Ubound(NewsArray)
	  'Arr_i=NewsNum_i
	  If CollecNewsAll>=CollecNewsNum And CollecNewsNum<>0 then Exit For'������
      CollecNewsAll=CollecNewsAll+1
      UploadFiles=""
      DefaultPicUrl=""
      IncludePic=0
      ImagesNum=0
      NewsCode=""
      FoundErr=False
      ErrMsg=""
      His_Repeat=False
	  NewsUrl=NewsArray(Arr_i)
      Title=""
      PaingNum=1
      If Response.IsClientConnected Then 
         Response.Flush 
      Else 
         Response.End 
      End If

      If CollecTest=False Then
         His_Repeat=SK.CheckRepeat(NewsUrl)
      Else
         His_Repeat=False
      End If
      If His_Repeat=True Then
         FoundErr=True
      End If
	  If FoundErr<>True then
		  If x_tp=1 then
		  	'On Error Resume Next
			picpath=Newsimage(arr_i)
			iF SaveFiles=1 then picpath=Skcj.Sk_SaveFile(Colleclx,picpath)
		  End if
	  End if
      If FoundErr<>True Then
         NewsCode=Skcj.ReplaceTrim(skcj.GetHttpPage(NewsUrl,selEncoding))
         If NewsCode="$False$" Then
            FoundErr=True
			ErrMsg=ErrMsg & "<br>�ڻ�ȡ��" & NewsUrl & "��ҳԴ��ʱ��������"
			Title="��ȡ��ҳԴ��ʧ��"
         End If
      End If

      If FoundErr<>True Then
         Title=Skcj.GetBody(NewsCode,TsString,ToString,False,False)
         If Title="$False$" or Title="" then
            FoundErr=True
            ErrMsg=ErrMsg & "<br>�ڷ�����" & NewsUrl & "��"& Skcj.GetItemConfig("CjName",Colleclx) &"����ʱ��������"
			Title="<br>�����������" 
         End If
         If FoundErr<>True Then
		 	if CsString<>"0" or CoString<>"0" then
            Content=Skcj.GetBody(NewsCode,CsString,CoString,False,False)
            else
			Content=""
			end if
			If Content="$False$" Then
               'FoundErr=True
               'ErrMsg=ErrMsg & "<br>�ڷ�����" & NewsUrl & "��"& Skcj.GetItemConfig("CjName",Colleclx) &"����ʱ��������"
               Title=Title & "<br>���ķ�������" 
            End If
         End If
         If FoundErr<>True Then
         If NewsPaingType=1 Then
            NewsPaingNext=Skcj.GetBody(NewsCode,NPsString,NPoString,False,False)
			   If NewsPaingNext<>"$False$" Then
			   		 NewsPaingNext_Code=Skcj.GetArray(NewsPaingNext,NewsUrlPaing_s,NewsUrlPaing_o,False,False)
					 TypeArray_Url=Split( NewsPaingNext_Code,"$Array$")
					 if Ubound(TypeArray_Url)<>0 Then
						 For i=0 to Ubound(TypeArray_Url)
							Call Sk.ShowMsg_1("��ҳ�ɼ���...  ��ǰ�ɼ���"&I+1&"ҳ<br>")
							Response.Flush()
							 TypeNews_Url=Trim(Skcj.FormatRemoteUrl(TypeArray_Url(i),NewsUrl))
							 NewsPaingNextCode=Skcj.ReplaceTrim(skcj.GetHttpPage(TypeNews_Url,selEncoding)) 
							 '---------------------------ͼƬ��ҳ--------------------------------------------
							 IF Colleclx=2 Then 
								 	PicUrls=Skcj.GetBody(NewsPaingNextCode,photourls,photourlo,False,False)
									PicUrls=Trim(Skcj.FormatRemoteUrl(PicUrls,NewsUrl))
							 		IF SaveFiles=1 then 
										PicUrls=Skcj.Sk_SaveFile(Colleclx,PicUrls)
										If PicUrls=False then
										Response.Write "&nbsp;----" & PicUrls & " ����ʧ��<br>"
										Else
										Response.Write "&nbsp;" & Skcj.GetItemConfig("CjName",Colleclx) & I &"--" & PicUrls & " ����ɹ�<br>"
										End if
										Response.Flush()
									End IF
									if PicUrls<>False then
										If i=0 then
												PicUrls_i="ͼƬ1|" & PicUrls 
										Else
												PicUrls_i= PicUrls_i & "|||ͼƬ" & i  & "|" & PicUrls 
										End if
									End if	
									PicUrls=PicUrls_i
							 End if
							 '---------------------------ͼƬ��ҳ------------------------------------------------
							 ContentTemp=Skcj.GetBody(NewsPaingNextCode,CsString,CoString,False,False)
							 If ContentTemp<>"$False$" then Content=Content & "[NextPage]" & ContentTemp
						 Next
					 End if
               End If
         End If
		 '����
		 SK.FilterScript()
		 SK.GetFilters
         SK.Filters
         Title=FpHtmlEnCode(Title)
         Content=Ubbcode(Content)
		 Content=Skcj.ItemReplaceStr(Content,strReplace)'�����滻
         End If
      End If
	  
If Colleclx=2 And FoundErr<>True then 'ͼƬ����
	'--------------------------------���3�ɼ�-------------------------------------	
	IF NewsPaingType=2 Then
			i=1
			ListTypeCode=Skcj.GetBody(NewsCode,PhotoType_s,PhotoType_o,False,False)
			If ListTypeCode<>"$False$" Then 
				ListTypeUrlCode=Skcj.GetArray(ListTypeCode,PhotoLurl_s,PhotoLurl_o,False,False)
				If Phototypefy_s<>"0" AND Phototypefy_o<>"0" AND Phototypefyurl_s<>"0" AND Phototypefyurl_o<>"0" Then
				ListTypeCode_2=Skcj.GetBody(NewsCode,Phototypefy_s,Phototypefy_o,False,False)
					If ListTypeCode_2<>"$False$"   Then 
						ListTypeUrlCode_2=Skcj.GetArray(ListTypeCode_2,Phototypefyurl_s,Phototypefyurl_o,False,False)
						TypeUrlArray_2=Split(ListTypeUrlCode_2,"$Array$")
						For Arr_ii_2=0 to Ubound(TypeUrlArray_2)
						TypeNewsUrl=Trim(Skcj.FormatRemoteUrl(TypeUrlArray_2(Arr_ii_2),NewsUrl))
							NewsTypeCode=Skcj.ReplaceTrim(skcj.GetHttpPage(TypeNewsUrl,selEncoding))
							ListTypeCode=Skcj.GetBody(NewsTypeCode,PhotoType_s,PhotoType_o,False,False)
							If ListTypeCode<>"$False$" Then 
								ListTypeUrlCode=Skcj.GetArray(ListTypeCode,PhotoLurl_s,PhotoLurl_o,False,False)
								TypeUrlArray=Split(ListTypeUrlCode,"$Array$")
								For Arr_ii=0 to Ubound(TypeUrlArray)
									TypeNewsUrl=Trim(Skcj.FormatRemoteUrl(TypeUrlArray(Arr_ii),NewsUrl))
									If TypeNewsUrl<>"$False$" Then 
										
										if Phototypeurl_s<>"0" or Phototypeurl_o<>"0" then
											NewsTypeCode=Skcj.ReplaceTrim(skcj.GetHttpPage(TypeNewsUrl,selEncoding))
											PicUrls=Skcj.GetBody(NewsTypeCode,Phototypeurl_s,Phototypeurl_o,False,False)
											PicUrls=Trim(Skcj.FormatRemoteUrl(PicUrls,TypeNewsUrl))
											if HttpUrlStr<>"" then PicUrls=HttpUrlStr & PicUrls'�ض���ַ
										Else
											PicUrls=TypeNewsUrl
										end if
										IF SaveFiles=1 then 
											PicUrls=Skcj.Sk_SaveFile(Colleclx,PicUrls)
											if PicUrls=False then
											Response.Write "&nbsp;----" & PicUrls & " ����ʧ��<br>"
											Else
											Response.Write "&nbsp;" & Skcj.GetItemConfig("CjName",Colleclx) & I &"--" & PicUrls & " ����ɹ�<br>"
											End if
											Response.Flush()
										End IF
										if PicUrls<>False then
											If arr_ii=0 and Arr_ii_2=0 then
													PicUrls_i="ͼƬ1|" & PicUrls 
													i=i+1
											Else
													PicUrls_i= PicUrls_i & "|||ͼƬ" & i  & "|" & PicUrls 
													i=i+1
											End if
											PicUrls=PicUrls_i
										End if
									End If
								Next
							End If
						Next
						PicUrls=PicUrls_i
						Call sk.SaveArticle
					Else
						Call sk.Coll_ListType_2	
					End if
				Else
					Call sk.Coll_ListType_2	
				End If
			Else
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & NewsUrl & "2�������б�Դ��ʱ��������</li>"	
			End If
	End if
	'-----------------------------���3�ɼ�----------------------------------------
	If NewsPaingType=0 Then
		If Downlist_s="" or  Downlist_o="" or DownUrl_s="" or DownUrl_o="" then'ͼƬ����
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>ͼƬ��ַ���ò���Ϊ��</li>" 
		Else	
			DownUrls=Skcj.GetBody(NewsCode,Downlist_s,Downlist_o,False,False)
			If DownUrls<>"$False$" then
				DownUrls=Skcj.GetBody(DownUrls,DownUrl_s,DownUrl_o,False,False)
				IF DownUrls<>"$False$" then	
					DownUrls=Trim(Skcj.FormatRemoteUrl(DownUrls,NewsUrl))
					IF SaveFiles=1 then
						DownUrls=Skcj.Sk_SaveFile(Colleclx,DownUrls)
						if DownUrls=False then
							Response.Write "&nbsp;----" & DownUrls & " ����ʧ��<br>"
						Else
							Response.Write "&nbsp;ͼƬ" & DownUrls & " ����ɹ�<br>"
						End if
						Response.Flush()
					End IF
					PicUrls=DownUrls
					PicUrls= "ͼƬ1|" & PicUrls
				Else
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & NewsUrl & "ͼƬ����Դ��ʱ��������</li>"
				End if
			Else
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & NewsUrl & "ͼƬ�б�Դ��ʱ��������</li>"
			End if
		End if
	End if
End If
	  	  
If Colleclx=3 And FoundErr<>True then '����
dim DownUrls_i
  	If Downlist_s="" or  Downlist_o="" or DownUrl_s="" or DownUrl_o="" then'���ص�ַ����
    	FoundErr=True
    	ErrMsg=ErrMsg & "<br><li>���ص�ַ���ò���Ϊ��</li>" 
	Else	
		DownUrls=Skcj.GetBody(NewsCode,Downlist_s,Downlist_o,False,False)
		If DownUrls<>"$False$" then
			IF LinkUrlYn=1 then
				DownUrls=Skcj.GetArray(DownUrls,DownUrl_s,DownUrl_o,False,False)
			else
				DownUrls=Skcj.GetBody(DownUrls,DownUrl_s,DownUrl_o,False,False)
			end if
			IF DownUrls<>"$False$" then
					if LinkUrlYn=1 then
						i=1	
						TypeUrlArray=Split(DownUrls,"$Array$")
						For Arr_ii=0 to Ubound(TypeUrlArray)
							DownUrls=Trim(Skcj.FormatRemoteUrl(TypeUrlArray(Arr_ii),NewsUrl))
							If arr_ii=0  then
								DownUrls_i="���ص�ַ1|" & DownUrls
								i=i+1
							Else
								DownUrls_i= DownUrls_i & "|||���ص�ַ" & i  & "|" & DownUrls
								i=i+1
							End if
						Next
						DownUrls=DownUrls_i
					Else
						DownUrls=Trim(Skcj.FormatRemoteUrl(DownUrls,NewsUrl))
						DownUrls="���ص�ַ1|" & DownUrls
					End if	 
			Else
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & NewsUrl & "���ص�ַ����Դ��ʱ��������</li>"
			End if
		Else
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & NewsUrl & "���ص�ַ�б�Դ��ʱ��������</li>"
		End if
	End if
	
	If DownNewType=1 then'�´��ڴ���������
		If DownNewlist_s<>"" or  DownNewlist_o<>"" or DownNewUrl_s<>"" or DownNewUrl_o<>"" then
			DownUrls=Replace(DownUrls,"���ص�ַ1|","")
			DownUrls=Skcj.ReplaceTrim(skcj.GetHttpPage(DownUrls,selEncoding))
			DownUrls=Skcj.GetBody(DownUrls,DownNewlist_s,DownNewlist_o,False,False)
			If DownUrls<>"$False$" then 
				DownUrls=Skcj.GetArray(DownUrls,DownNewUrl_s,DownNewUrl_o,False,False)
				IF DownUrls<>"$False$" then
					i=1	
					TypeUrlArray=Split(DownUrls,"$Array$")
					For Arr_ii=0 to Ubound(TypeUrlArray)
						DownUrls=Trim(Skcj.FormatRemoteUrl(TypeUrlArray(Arr_ii),NewsUrl))
						If arr_ii=0  then
							DownUrls_i="���ص�ַ1|" & DownUrls
							i=i+1
						Else
							DownUrls_i= DownUrls_i & "|||���ص�ַ" & i  & "|" & DownUrls
							i=i+1
						End if
					Next
					DownUrls=DownUrls_i
				Else
					'FoundErr=True
    				ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & NewsUrl & "(�´���)���ص�ַ�б�Դ��ʱ��������</li>" 
				End if
			Else
				'FoundErr=True
    			ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & NewsUrl & "(�´���)���ص�ַ�б�Դ��ʱ��������</li>" 
			End if
		Else
			FoundErr=True
    		ErrMsg=ErrMsg & "<br><li>���ص�ַ���ò���Ϊ��</li>" 
		End if
	End if 
	
	If ZdType_001=0 then'������С����
   		DownSize=""
	Else
		If  Zds_001="0" and Zdo_001<>"" then
			DownSize=Zdo_001
		Else
			DownSize=FpHtmlEncode(Skcj.GetBody(NewsCode,Zds_001,Zdo_001,False,False))
		End If
	End If
	If ZdType_002=0 then'������������
   		DownYY=""
	Else
		If  Zds_002="0" and Zdo_002<>"" then
			DownYY=Zdo_002
		Else
			DownYY=FpHtmlEncode(Skcj.GetBody(NewsCode,Zds_002,Zdo_002,False,False))
		End If
	End If
 	If ZdType_003=0 then'��Ȩ��ʽ����
   		DownSQ=""
	Else
		If  Zds_003="0" and Zdo_003<>"" then
			DownSQ=Zdo_003
		Else
			DownSQ=FpHtmlEncode(Skcj.GetBody(NewsCode,Zds_003,Zdo_003,False,False))
		End If
	End If
 	If ZdType_004=0 then'���л�������
   		DownPT=""
	Else
		If  Zds_004="0" and Zdo_004<>"" then
			DownPT=Zdo_004
		Else
			DownPT=FpHtmlEncode(Skcj.GetBody(NewsCode,Zds_004,Zdo_004,False,False))
		End If
	End If
 	If ZdType_005=0 then'��ʾ��ַ����
   		YSDZ=""
	Else
		If  Zds_005="0" and Zdo_005<>"" then
			YSDZ=Zdo_005
		Else
			YSDZ=FpHtmlEncode(Skcj.GetBody(NewsCode,Zds_005,Zdo_005,False,False))
		End If
	End If
 	If ZdType_006=0 then'ע���ַ����
   		ZCDZ=""
	Else
		If  Zds_006="0" and Zdo_006<>"" then
			ZCDZ=Zdo_006
		Else
			ZCDZ=FpHtmlEncode(Skcj.GetBody(NewsCode,Zds_006,Zdo_006,False,False))
		End If
	End If	
	
 	If ZdType_007=0 then'����ͼƬ����
   		PhotoUrl=""
	Else
		If  Zds_007="0" and Zdo_007<>"" then
			PhotoUrl=Zdo_007
		Else
			PhotoUrl=Skcj.GetBody(NewsCode,Zds_007,Zdo_007,False,False)
			PhotoUrl=Trim(Skcj.FormatRemoteUrl(PhotoUrl,NewsUrl))
		End If
	End If  
End if

IF Colleclx=4 And FoundErr<>True then '����
	If Downlist_s="" or  Downlist_o="" or DownUrl_s="" or DownUrl_o="" then'���ص�ַ����
    	FoundErr=True
    	ErrMsg=ErrMsg & "<br><li>�������ص�ַ���ò���Ϊ��</li>" 
	Else	
		DownUrls=Skcj.GetBody(NewsCode,Downlist_s,Downlist_o,False,False)
		If DownUrls<>"$False$" then
			DownUrls=Skcj.GetBody(DownUrls,DownUrl_s,DownUrl_o,False,False)
			IF DownUrls<>"$False$" then
						DownUrls=Trim(Skcj.FormatRemoteUrl(DownUrls,NewsUrl))
						IF SaveFiles=1 then 
							DownUrls=Skcj.Sk_SaveFile(Colleclx,DownUrls)
						End IF 
			Else
				'FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & NewsUrl & "���ص�ַ����Դ��ʱ��������</li>"
			End if
		Else
			'FoundErr=True
			ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & NewsUrl & "���ص�ַ�б�Դ��ʱ��������</li>"
		End if
	End if
	
	If DownNewType=1 then'�´��ڴ���������
		If DownNewlist_s<>"" or  DownNewlist_o<>"" or DownNewUrl_s<>"" or DownNewUrl_o<>"" then
			DownUrls=Skcj.ReplaceTrim(skcj.GetHttpPage(DownUrls,selEncoding))
			DownUrls=Skcj.GetBody(DownUrls,DownNewlist_s,DownNewlist_o,False,False)
			If DownUrls<>"$False$" then 
				DownUrls=Skcj.GetBody(DownUrls,DownNewUrl_s,DownNewUrl_o,False,False)
				IF DownUrls<>"$False$" then
					DownUrls=Trim(Skcj.FormatRemoteUrl(DownUrls,NewsUrl))
				Else
					'FoundErr=True
    				ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & NewsUrl & "(�´���)���ص�ַ�б�Դ��ʱ��������</li>" 
				End if
			Else
				'FoundErr=True
    			ErrMsg=ErrMsg & "<br><li>�ڻ�ȡ:" & NewsUrl & "(�´���)���ص�ַ�б�Դ��ʱ��������</li>" 
			End if
		Else
			FoundErr=True
    		ErrMsg=ErrMsg & "<br><li>���ص�ַ���ò���Ϊ��</li>" 
		End if
	End if 
End if 	  


	  
      If FoundErr<>True Then
         If DateType=0 Then
            UpDateTime=Now()
         Else
		 	If DateType=1 Then
               UpDateTime=Skcj.GetBody(NewsCode,DsString,DoString,False,False)
               UpDateTime=FpHtmlEncode(UpDateTime)
               UpDateTime=Trim(Replace(UpDateTime,"&nbsp;"," "))
               If IsDate(UpDateTime)=True Then
                  UpDateTime=CDate(UpDateTime)
               Else
                  UpDateTime=Now()
               End If
            End If
         End If
                  
         If AuthorType=1 Then
            Author=Skcj.GetBody(NewsCode,AsString,AoString,False,False)
         ElseIf AuthorType=2 Then
            Author=AuthorStr
         Else
            Author="����"
         End If
         Author=FpHtmlEncode(Author)
         If Author="" or Author="$False$" then
            Author="����"
         Else
            If Len(Author)>255 then
               Author=Left(Author,255)
            End If
         End If
           
         If CopyFromType=1 Then
            CopyFrom=Skcj.GetBody(NewsCode,FsString,FoString,False,False)
         ElseIf CopyFromType=2 Then
            CopyFrom=CopyFromStr
         Else
            CopyFrom="����"
         End If
         CopyFrom=FpHtmlEncode(CopyFrom)
         If CopyFrom="" or CopyFrom="$False$" Then
	            CopyFrom="����"
         Else
            If Len(CopyFrom)>255 Then 
               CopyFrom=Left(CopyFrom,255)
            End If
         End If
         If KeyType=0 Then
            Key=Title
            Key=CreateKeyWord(Key,2)
         ElseIf KeyType=1 Then
            Key=Skcj.GetBody(NewsCode,KsString,KoString,False,False)
			Key=Replace(Key,",","|")
			Key=Replace(Key," ","|")
            Key=FpHtmlEncode(Key)
            'Key=CreateKeyWord(Key,2)
         ElseIf KeyType=2 Then
            Key=KeyStr
            Key=FpHtmlEncode(Key)
            If Len(Key)>253 Then
               Key= Left(Key,253)
            Else
               Key=Key
            End If
         End If
         If Key="" or Key="$False$" Then
            Key=KeyStr
         End If
		 
         If UploadFiles<>"" Then
            If Instr(UploadFiles,"|")>0 Then
               Arr_Images=Split(UploadFiles,"|") 
               ImagesNum=Ubound(Arr_Images)+1
               DefaultPicUrl=Arr_Images(0)
            Else
               ImagesNum=1
               DefaultPicUrl=UploadFiles
            End If

            If DefaultPicYn=False then
               DefaultPicUrl=""
            End If
            If IncludePicYn=True Then
               IncludePic=-1
            Else
               IncludePic=0
            End If
            If SaveFiles<>True Then
               UploadFiles=""
            End If
         Else
            ImagesNum=0
            DefaultPicUrl=""
            IncludePic=0         
         End If
         ImagesNumAll=ImagesNumAll+ImagesNum
      End If

	  if  Colleclx =1 then 
	  	set rs = ConnItem.execute("select top 1 Dir from SK_cj where ID="& Colleclx)
			IF SaveFiles=1 then
				Content=Skcj.ReplaceSaveRemoteFile(Content,"/",rs("Dir") & SaveFileUrl,True,NewsUrl)'Զ��ͼƬ
				Content=SKcj.ReSaveRemoteFile(Content,NewsUrl,rs("Dir") & SaveFileUrl,True)'Զ���ļ�
			Else
				Content=Skcj.ReplaceSaveRemoteFile(Content,"/",rs("Dir"),False,NewsUrl)'Զ��ͼƬ
				Content=Skcj.ReSaveRemoteFile(Content,NewsUrl,rs("Dir"),False)'Զ���ļ�
			End if 
		rs.close
		set rs=nothing
	  End if
	  '--
      If FoundErr<>True Then
			If His_Repeat<>True Then
				Call sk.SaveArticle
			End if		
	  	If CollecTest=False Then
            SqlItem="INSERT INTO Histroly(ItemID,ChannelID,ClassID,SpecialID,ArticleID,Title,CollecDate,NewsUrl,Result) VALUES ('" & ItemID & "','" & ChannelID & "','" & ClassID & "','" & SpecialID & "','" & ArticleID & "','" & Title & "','" & Now() & "','" & NewsUrl & "',True)"
            ConnItem.Execute(SqlItem)
            Content=Replace(Content,"[InstallDir_ChannelDir]",strInstallDir & strChannelDir & "/")
         End If
         NewsSuccesNum=NewsSuccesNum+1
         ErrMsg=ErrMsg & "No:<font color=red>" & NewsSuccesNum+NewsFalseNum & "</font><br>"
		 ErrMsg=ErrMsg & Skcj.GetItemConfig("CjName",Colleclx) &"���⣺"
         ErrMsg=ErrMsg & "<font color=red>" & Title & "</font><br>"
         ErrMsg=ErrMsg & "����ʱ�䣺" & UpDateTime & "<br>"
		 If Colleclx=1 then ErrMsg=ErrMsg & "���ű��⣺" : ErrMsg=ErrMsg & "�������ߣ�" & Author & "<br>" : ErrMsg=ErrMsg & "������Դ��" & CopyFrom & "<br>"
         ErrMsg=ErrMsg & "�ɼ�ҳ�棺<a href=" & NewsUrl & " target=_blank>" & NewsUrl & "</a><br>"
		 if x_tp =1 then ErrMsg=ErrMsg & "�ɼ�СͼƬ��<a href=" & picpath & " target=_blank>" & picpath & "</a><br>"
		 
		 ErrMsg=ErrMsg & "����Ԥ����"
         If Content_View=True Then
            ErrMsg=ErrMsg & "<br>" & Content
         Else
            ErrMsg=ErrMsg & "��û����������Ԥ������"
         End If
         ErrMsg=ErrMsg & "<br><br>�� �� �֣�" & Key & ""
      Else
         NewsFalseNum=NewsFalseNum+1
         If His_Repeat=True Then
            ErrMsg=ErrMsg & "No:<font color=red>" & NewsSuccesNum+NewsFalseNum & "</font><br>"
			ErrMsg=ErrMsg & "Ŀ��"& Skcj.GetItemConfig("CjName",Colleclx) &"��<font color=red>"
            If His_Result=True Then
               ErrMsg=ErrMsg & His_Title
            Else
               ErrMsg=ErrMsg & NewsUrl
            End If
            ErrMsg=ErrMsg & "</font> �ļ�¼�Ѵ��ڣ�������ɼ���<br>"
            ErrMsg=ErrMsg & "�ɼ�ʱ�䣺" & His_CollecDate & "<br>"
			ErrMsg=ErrMsg & ""& Skcj.GetItemConfig("CjName",Colleclx) &"��Դ��<a href='" & NewsUrl & "' target=_blank>"&NewsUrl&"</a><br>"
            ErrMsg=ErrMsg & "�ɼ������"
            If His_Result=False Then
               ErrMsg=ErrMsg & "ʧ��"
               ErrMsg=ErrMsg & "<br>ʧ��ԭ��" & Title
            Else
               ErrMsg=ErrMsg & "�ɹ�"
            End If            
            ErrMsg=ErrMsg & "<br>��ʾ��Ϣ�������ٴβɼ������Ƚ������ŵ���ʷ��¼<font color=red>ɾ��</font><br>"
         End If
         If CollecTest=False And His_Repeat=False Then
            SqlItem="INSERT INTO Histroly(ItemID,ChannelID,ClassID,SpecialID,Title,CollecDate,NewsUrl,Result) VALUES ('" & ItemID & "','" & ChannelID & "','" & ClassID & "','" & SpecialID & "','" & Title & "','" & Now() & "','" & NewsUrl & "',False)"
            ConnItem.Execute(SqlItem)
         End If
	  	 
      End If
	   Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""2"" class=""tableBorder""  cellspacing=""1"">"
	   Response.Write "   <tr class='tdbg'>"          
	   Response.Write "      <td height=""22"" colspan=""2"" align=""left"">"
	   Response.Write ErrMsg
	   Response.Write "      </td>"
	   Response.Write "   </tr><br>"
	   Response.Write "</table>"
	   Response.Flush()'ˢ�� 
   Next    
			If ListEnd<>true then
				if Collecdate<>"" Then 
					Collecdate=Day(now())
					response.write("<script>location.href='sk_Timing.asp?action=GoTiming&Collecdate="&  Day(now()) &"';</script>")'��ҳ��
				Else
					response.write("<script>location.href='Sk_CollectionFast.asp?ItemID="& ItemID &"&ItemNum=" & ItemNum & "&ListNum=" & ListNum +1 & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ListPaingNext=" & ListPaingNext &"&NewsNum_i="& 0 &"&Itemok="& Itemok &"&Itemon="& Itemon &"&Collecdate="& Collecdate &"';</script>")'���
				End if	
			End if
Else
   If FoundErr_1=True Then
		response.write("<script>location.href='Sk_CollectionFast.asp?ItemID="& ItemID &"&ItemNum=" & ItemNum & "&ListNum=" & ListNum +1 & "&NewsSuccesNum=" & NewsSuccesNum & "&NewsFalseNum=" & NewsFalseNum & "&ImagesNumAll=" & ImagesNumAll & "&ListPaingNext=" & ListPaingNext &"&NewsNum_i="& 0 &"&Itemok="& Itemok &"&Itemon="& Itemon &"&Collecdate="& Collecdate &"';</script>")'���
	 	FoundErr_1=False
	End If
   Call Sk.ShowMsg(ErrMsg)
End If
End Sub
%>