<%
'************************************************************ 
'���ߣ�yujianyue (12391.net) 
'��Ȩ��Դ���빫�����ڱ�������Ȩǰ���£�������;�������ʹ�á� 
'������2018-11-08 
'��ϵ��15058593138@qq.com 
'��ע��������Ѱ汾���ɱ�ѡ��������ʣ�http://12391.net
'************************************************************
%><!--#include file="inc/conn.asp"--><!--#include file="inc/pubs.asp"-->
<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="gb2312" />
<meta name="viewport" content="width=device-width,minimum-scale=1.0,maximum-scale=1.0" />
<meta name="apple-mobile-web-app-capable" content="yes" />
<title><%=title%>,96448.cn</title>
<!--������������뱣������author copyright���� -->
<meta name="author" content="yujianyue, admin@ewuyi.net">
<meta name="copyright" content="www.12391.net">
<!--������������뱣������author copyright���� -->
<script type="text/javascript" src="inc/ajax.js"></script>
<link href="inc/web.css" rel="stylesheet" type="text/css" />
<body onLoad="inst();">
<div class="html">
<div class="sub_bod"></div>
<div class="divs" id="divs">
<div id="head" class="head" onclick="location.href='?t=<%=tima%>';">
<%=title%>
</div>
<div class="main" id="main">
<%
startime=timer()
duansss="������"&duan

mamasss="������4������֤��"

names=""&trim(request("name"))&""
codes=""&trim(request("code"))&""

if names="" then
%>
<form name="queryForm" method="post" class="" action="" onsubmit="return startRequest(0);">
<div class="so_box" id="11">
<input name="name" type="text" class="txts" id="name" value="<%=duansss%>" placeholder="<%=duansss%>" onfocus="st('name',1)" onBlur="startRequest(2)" />
</div>
<%if yanzhenma="1" then
mamas="+��֤��"%>
<div class="so_box" id="33">
<input name="code" type="text" class="txts" id="code" value="<%=mamasss%>" placeholder="<%=mamasss%>" onfocus="this.value=''" onBlur="startRequest(3)" />
<div class="more" id="clearkey">
<img src="inc/Code.asp?t=<%=timer%>" id="Codes" onClick="this.src='inc/Code.asp?t='+new Date();"/>
</div></div><%end if%>
<div class="so_but">
<input type="submit" name="button" class="buts" id="sub" value="������ѯ" />
<input type="button" class="buts" value="ˢ�±�ҳ" name="print" onclick="location.reload();">
</div>
<div class="so_bus" id="tisha">
<strong>˵��</strong>:<%=duan%><%=mamas%>��������ȷ����ʾ��Ӧ�����
<br><!--��������ֶ�����html˵����ʼ�����������뺺�ּ��ɡ�-->
<%
descfile=replace(descfile,"../","")
if IsFile(descfile)=True then
dushuoming = FsoFileRead(descfile,"gb2312")
else
dushuoming = "<!--û�ж�ȡ������-->"
end if
  Response.Write dushuoming
%>
<!--��������ֶ�����html˵�����������������뺺�ּ��ɡ�-->
</div>
<div id="tishi1" style="display:none;"><%=duansss%></div>
<div id="tishi4" style="display:none;"><%=mamasss%></div>
</form>
<%
else

datas=""&UpDir&"/"&times&""&mdbtype

if yanzhenma="1" then
if len(codes)<>4 or codes<>Session("GetCode") Then
 call AlertBack("��������ȷ����֤��Ŷ��") 
End if
end if

if len(names)<1 and len(names)>18 Then
 call AlertBack("������"&duan&"��") 
End if

if IsFile(datas)=True then
else
 call AlertBack("������ʱû���ϴ����߲�����Ŷ��"&datas) 
end if


page=request("p")
if isnumeric(page)=false or len(page)=0 then
page="1"
end if

thekey=names
thekey=left(thekey,8)
if len(thekey)=0 then
thekeys=""
else
thekeys="&q="&thekey
end if

'�������ݿ��ܿ�
set conn=Server.CreateObject("ADODB.Connection")
conn.open "DRIVER=Driver do Microsoft Access (*.mdb);UID=admin;PWD=;DBQ="&Server.MapPath(""&datas&"")
If Err Then
err.Clear
response.write "���ݿ����ӳ������������ִ�(1)"&vbcrlf
response.End()
End If
Response.Write "<h1>&nbsp;</h1><!--startprint-->"&vbcrlf
Const adSchemaTables = 20
set objSchema = conn.OpenSchema(adSchemaTables)
     bnames="---"
Do While Not objSchema.EOF
	if objSchema("TABLE_TYPE") = "TABLE" then
        tablename = objSchema("TABLE_NAME")
      bnames=bnames&tablename&"---"
	end if
objSchema.MoveNext
Loop

if instr(bnames,"---"&biaogege1&"---")>0 then
else
 biaogege1 = split(bnames,"---")(1) 
end if
if instr(bnames,"---"&biaogege1&"---")>0 then
else
 call AlertBack("���������õ����ݱ���["&biaogege1&"]�Ƿ���ڣ�") 
end if

Response.Write "<table cellspacing=""0"">"&vbcrlf
Response.Write "<caption align='center'>"&times&" ��ѯ���</caption>"&vbcrlf

set rsdo=Server.CreateObject("ADODB.RecordSet")
if len(thekey)>0 and thekey<>"���������������ѯ" then
sqldo="select * from ["&biaogege1&"] where ["&tiaojian1&"] = '"&thekey&"' "
else
sqldo="select * from ["&biaogege1&"] "
end if
'sqldo=sqldo&" order by id desc" '�̶������㷨
rsdo.open sqldo,conn,1,1
rsdo.PageSize=20
lies = rsdo.fields.count
     tnames="---"
   response.write "<tr class=""tt"">"&vbcrlf
   for i = 0 to lies - 1         'ѭ���ֶ���
      lieti = rsdo.fields.item(i).name
   response.write  "<td>" & lieti & " </td>"&vbcrlf
      tnames=tnames&lieti&"---"
   next
   response.write "</tr>"&vbcrlf

if instr(tnames,"---"&tiaojian1&"---")>0 then
else
'call AlertBack("���������õĲ�ѯ����["&tiaojian1&"]�Ƿ���ڣ�") 
end if

If not (rsdo.bof and rsdo.eof) then
 rsdo.AbsolutePage=page
 for k=1 to rsdo.PageSize

   response.write "<tr>"&vbcrlf
      for i = 0 to lies - 1
        curValue = rsdo.fields.item(i).value
 If IsNull(curValue) or len(curValue)<1 Then
  curValue="&nbsp;"
 End If
    response.write "<td>" & curValue & "</td>"&vbcrlf
      next
    response.write "</tr>"&vbcrlf

 rsdo.movenext
 If rsdo.EOF Then Exit For
 next
response.write "<tr><td colspan="""&lies&""" class=""titi"">"
rc=rsdo.RecordCount
ps=rsdo.PageSize
pc=rsdo.PageCount
'response.write getPage(page,pc)
rsdo.close
set rsdo=nothing
response.write "</td></tr>"
else
response.write "<tr><td colspan="""&lies&""">"
response.write "<p>��û��ѯ��Ϣ��</p>"
response.write "</td></tr>"
end if

Response.Write "</table>"&vbcrlf
Response.Write "<!--endprint-->"&vbcrlf

%>
<div class="so_but">
<input type="button" class="buts" value="Ԥ ��" name="print" onclick="preview()">
<input type="button" class="buts" value="�� ��" id="reset" onclick="location.href='?b=<%=tima%>';">
</div>
<%end if
endtime=timer()%>
</div>
<div class="boto" id="boto">
&copy; <%=year(now)%>&nbsp; <a href="<%=copysu%>" target="_blank"><%=copysr%></a>
<!--����Դ�����ߣ��뱣������һ�У�����ʾ�ģ��ڲ�ѯҳ��-->
<!-- ������Ѱ汾���ɱ�ѡ��������ʣ�http://12391.net -->
<!--����Դ�����ߣ��뱣������һ�У�����ʾ�ģ��ڲ�ѯҳ��-->
</div>
</div>
</div>
</body>
</html>