<%
'=====================================================================
' �������ƣ��ɼ�ϵͳ
' �ٷ���վ��(www.iising.com) ����QQ��24387481
' ���÷������˻��漼�����Ժܺý���ⷽ���ì�ܣ�������ASP���к͸���Ч�ʡ� 
'===================================================================== 
class SK_clsCache
'----------------------------
private cache           '��������
private cacheName       '����Application����
private expireTime      '�������ʱ��
private expireTimeName  '�������ʱ��Application����
private path            '����ҳURL·��
private vaild           'ansir����
private sub class_initialize()
path=request.servervariables("url")
path=left(path,instrRev(path,"/"))
end sub

private sub class_terminate()
end sub

Public Property Get Version
	Version="������2006"
End Property

public property get valid '��ȡ�����Ƿ���Ч/����
	if isempty(cache) or (not isdate(expireTime)) then
	vaild=false
	else
	valid=true
	end if
	end property
	
public property get value '��ȡ��ǰ��������/����
	if isempty(cache) or (not isDate(expireTime)) then
	value=null
	elseif CDate(expireTime)<now then
	value=null
	else
	value=cache
	end if
end property

public property let name(str) '���û�������/����
	cacheName=str&path
	cache=application(cacheName)
	expireTimeName=str&"expire"&path
	expireTime=application(expireTimeName)
end property

public property let expire(tm) '���û������ʱ��/����
	expireTime=tm
	application.Lock()
	application(expireTimeName)=expireTime
	application.UnLock()
end property

public sub add(varCache,varExpireTime) '�Ի��渳ֵ/����
	if isempty(varCache) or not isDate(varExpireTime) then
	exit sub
	end if
	cache=varCache
	expireTime=varExpireTime
	application.lock
	application(cacheName)=cache
	application(expireTimeName)=expireTime
	application.unlock
end sub

public sub clean() '�ͷŻ���/����
	application.lock
	application(cacheName)=empty
	application(expireTimeName)=empty
	application.unlock
	cache=empty
	expireTime=empty
end sub
 
public function verify(varcache2) '�Ƚϻ���ֵ�Ƿ���ͬ/�������������ǻ��
	if typename(cache)<>typename(varcache2) then
		verify=false
	elseif typename(cache)="Object" then
		if cache is varcache2 then
			verify=true
		else
			verify=false
		end if
	elseif typename(cache)="Variant()" then
		if join(cache,"^")=join(varcache2,"^") then
			verify=true
		else
			verify=false
		end if
	else
		if cache=varcache2 then
			verify=true
		else
			verify=false
		end if
	end if
end function
end class
%>