<%

'************************************************************ 
'���ߣ�yujianyue (12391.net) 
'��Ȩ��Դ���빫�����ڱ�������Ȩǰ���£�������;�������ʹ�á� 
'������2018-11-08 ������Ѱ汾���ɱ�ѡ��������ʣ�http://12391.net
'��ϵ��15058593138@qq.com 
'************************************************************


on error resume next

duan=tiaojian1

mdbtype=moexe 'ֻ����xls��ʽ�ļ�Ŷ��Ҫ�޸�

if len(tiaojian1)=2 then
 qianmian1=left(tiaojian1,1)&"&nbsp;&nbsp;"&right(tiaojian1,1)
else
 qianmian1=tiaojian1
end if


Function IsFile(FilePath)
 Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
 If (Fso.FileExists(Server.MapPath(FilePath))) Then
 IsFile=True
 Else
 IsFile=False
 End If
 Set Fso=Nothing
End Function

tima = time()

	'==============================
	'�� �� ����FsoFileRead
	'��    �ã���ȡ�ļ�
	'��    �����ļ����·��FilePath
	'==============================
	Function FsoFileRead(FilePath,charset)
	Set objAdoStream = Server.CreateObject("A"&"dod"&"b.St"&"r"&"eam")
	objAdoStream.Type=2
	objAdoStream.mode=3  
	objAdoStream.charset=charset
	objAdoStream.open 
	objAdoStream.LoadFromFile Server.MapPath(FilePath) 
	FsoFileRead=objAdoStream.ReadText 
	objAdoStream.Close
	Set objAdoStream=Nothing
	End Function
	

Function filemima(mimafile)
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set fd = fso.OpenTextFile(server.MapPath(mimafile), 1, True)
 if fd.AtEndOfStream=false then
 contentd = fd.readline()
 end if
filemima=trim(contentd)
 fd.close
end Function

Function filekey(texts)
 filekey=0
 rekey="-/-\-%-@-.-"
 keyes=split(rekey,"-")
 nnnnn=Ubound(keyes)
 For m=1 To Ubound(keyes)-1
 rekeys=keyes(m)
 rekeys=trim(rekeys)
 if instr(texts,rekeys)>0 and len(rekeys)>0 then
 filekey=filekey+1
 end if
 next
End Function

'==============================
'�� �� ���� AlertUrl(AlertStr,Url) 
'�� �ã������ת��ָ��ҳ��
'==============================
Function AlertUrl(AlertStr,Url) 
 Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />" &vbcrlf
 Response.Write "<script>" &vbcrlf
 Response.Write "alert('"&AlertStr&"');" &vbcrlf
 Response.Write "location.href='"&Url&"';" &vbcrlf
 Response.Write "</script>" &vbcrlf
 Response.End()
End Function
'==============================
'�� �� ���� AlertBack(AlertStr)
'�� �ã�����󷵻���һҳ��
'==============================
Function AlertBack(AlertStr) 
 Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />" &vbcrlf
 Response.Write "<script>" &vbcrlf
 Response.Write "alert('"&AlertStr&"');" &vbcrlf
 Response.Write "history.go(-1)" &vbcrlf
 Response.Write "</script>"&vbcrlf
 Response.End()
End Function

%>