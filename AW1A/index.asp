<%
'************************************************************ 
'作者：yujianyue (12391.net) 
'版权：源代码公开，在保留署名权前提下，各种用途均可免费使用。 
'创建：2018-11-08 
'联系：15058593138@qq.com 
'备注：更多免费版本今后可备选，详情访问：http://12391.net
'************************************************************
%><!--#include file="inc/conn.asp"--><!--#include file="inc/pubs.asp"-->
<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="gb2312" />
<meta name="viewport" content="width=device-width,minimum-scale=1.0,maximum-scale=1.0" />
<meta name="apple-mobile-web-app-capable" content="yes" />
<title><%=title%>,96448.cn</title>
<!--如果保留界面请保留以下author copyright两行 -->
<meta name="author" content="yujianyue, admin@ewuyi.net">
<meta name="copyright" content="www.12391.net">
<!--如果保留界面请保留以上author copyright两行 -->
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
duansss="请输入"&duan

mamasss="请输入4数字验证码"

names=""&trim(request("name"))&""
codes=""&trim(request("code"))&""

if names="" then
%>
<form name="queryForm" method="post" class="" action="" onsubmit="return startRequest(0);">
<div class="so_box" id="11">
<input name="name" type="text" class="txts" id="name" value="<%=duansss%>" placeholder="<%=duansss%>" onfocus="st('name',1)" onBlur="startRequest(2)" />
</div>
<%if yanzhenma="1" then
mamas="+验证码"%>
<div class="so_box" id="33">
<input name="code" type="text" class="txts" id="code" value="<%=mamasss%>" placeholder="<%=mamasss%>" onfocus="this.value=''" onBlur="startRequest(3)" />
<div class="more" id="clearkey">
<img src="inc/Code.asp?t=<%=timer%>" id="Codes" onClick="this.src='inc/Code.asp?t='+new Date();"/>
</div></div><%end if%>
<div class="so_but">
<input type="submit" name="button" class="buts" id="sub" value="立即查询" />
<input type="button" class="buts" value="刷新本页" name="print" onclick="location.reload();">
</div>
<div class="so_bus" id="tisha">
<strong>说明</strong>:<%=duan%><%=mamas%>都输入正确才显示相应结果。
<br><!--这里可以手动加入html说明开始，不懂的输入汉字即可。-->
<%
descfile=replace(descfile,"../","")
if IsFile(descfile)=True then
dushuoming = FsoFileRead(descfile,"gb2312")
else
dushuoming = "<!--没有读取到内容-->"
end if
  Response.Write dushuoming
%>
<!--这里可以手动加入html说明结束，不懂的输入汉字即可。-->
</div>
<div id="tishi1" style="display:none;"><%=duansss%></div>
<div id="tishi4" style="display:none;"><%=mamasss%></div>
</form>
<%
else

datas=""&UpDir&"/"&times&""&mdbtype

if yanzhenma="1" then
if len(codes)<>4 or codes<>Session("GetCode") Then
 call AlertBack("请输入正确的验证码哦！") 
End if
end if

if len(names)<1 and len(names)>18 Then
 call AlertBack("请输入"&duan&"！") 
End if

if IsFile(datas)=True then
else
 call AlertBack("数据暂时没有上传或者不存在哦！"&datas) 
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

'链接数据库总库
set conn=Server.CreateObject("ADODB.Connection")
conn.open "DRIVER=Driver do Microsoft Access (*.mdb);UID=admin;PWD=;DBQ="&Server.MapPath(""&datas&"")
If Err Then
err.Clear
response.write "数据库连接出错，请检查连接字串(1)"&vbcrlf
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
 call AlertBack("请检查你设置的数据表名["&biaogege1&"]是否存在！") 
end if

Response.Write "<table cellspacing=""0"">"&vbcrlf
Response.Write "<caption align='center'>"&times&" 查询结果</caption>"&vbcrlf

set rsdo=Server.CreateObject("ADODB.RecordSet")
if len(thekey)>0 and thekey<>"请输入标题域名查询" then
sqldo="select * from ["&biaogege1&"] where ["&tiaojian1&"] = '"&thekey&"' "
else
sqldo="select * from ["&biaogege1&"] "
end if
'sqldo=sqldo&" order by id desc" '固定排序算法
rsdo.open sqldo,conn,1,1
rsdo.PageSize=20
lies = rsdo.fields.count
     tnames="---"
   response.write "<tr class=""tt"">"&vbcrlf
   for i = 0 to lies - 1         '循环字段名
      lieti = rsdo.fields.item(i).name
   response.write  "<td>" & lieti & " </td>"&vbcrlf
      tnames=tnames&lieti&"---"
   next
   response.write "</tr>"&vbcrlf

if instr(tnames,"---"&tiaojian1&"---")>0 then
else
'call AlertBack("请检查你设置的查询条件["&tiaojian1&"]是否存在！") 
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
response.write "<p>暂没查询信息！</p>"
response.write "</td></tr>"
end if

Response.Write "</table>"&vbcrlf
Response.Write "<!--endprint-->"&vbcrlf

%>
<div class="so_but">
<input type="button" class="buts" value="预 览" name="print" onclick="preview()">
<input type="button" class="buts" value="返 回" id="reset" onclick="location.href='?b=<%=tima%>';">
</div>
<%end if
endtime=timer()%>
</div>
<div class="boto" id="boto">
&copy; <%=year(now)%>&nbsp; <a href="<%=copysu%>" target="_blank"><%=copysr%></a>
<!--尊重源码作者，请保留以下一行（不显示的）在查询页面-->
<!-- 更多免费版本今后可备选，详情访问：http://12391.net -->
<!--尊重源码作者，请保留以上一行（不显示的）在查询页面-->
</div>
</div>
</div>
</body>
</html>