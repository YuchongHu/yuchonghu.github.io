<%
cnt=0
dnt=0
s=Request.ServerVariables("path_translated")
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
cName=Server.MapPath("c.cnt")
dName=Server.MapPath("d.cnt")
Set objCountFile = objFSO.OpenTextFile(cName,1,True)
If Not objCountFile.AtEndOfStream Then cnt=CLng(objCountFile.ReadAll)
objCountFile.Close
Set objCountFile=Nothing
cnt=cnt+1
Set objCountFile=objFSO.CreateTextFile(cName,True)
objCountFile.Write cnt
objCountFile.Close
Set objCountFile=Nothing

if application("dntime")<=cint(hour(time())) then
 Set objCountFile = objFSO.OpenTextFile(dName,1,True)
 If Not objCountFile.AtEndOfStream Then dnt=CLng(objCountFile.ReadAll)
 objCountFile.Close
 Set objCountFile=Nothing
end if
application("dntime")=cint(hour(time()))
dnt=dnt+1
Set objCountFile=objFSO.CreateTextFile(dName,True)
objCountFile.Write dnt
objCountFile.Close
Set objCountFile=Nothing
Set objFSO = Nothing

t=(cint(day(date()))*24+cint(hour(time())))*60+cint(minute(time()))
k=0
i=1
y=0
Do While application("zxip"&i)<>""
 if application("zxip"&i)=Request.ServerVariables("REMOTE_ADDR") then
  application("zxsj"&i)=t
  y=1
 end if
 if t-application("zxsj"&i)>9 or t<application("zxsj"&i) then
  k=k+1
 else
  if k>0 then
   application.lock
   application("zxip"&i-k)=application("zxip"&i)
   application("zxsj"&i-k)=application("zxsj"&i)
   application.unlock
  end if
 end if
 if k>0 then
  application.lock
  application("zxip"&i)=""
  application.unlock
 end if
 i=i+1
loop
if y=0 then
 application("zxip"&i)=Request.ServerVariables("REMOTE_ADDR")
 application("zxsj"&i)=t
else
 i=i-1
end if%>
<HTML><HEAD><TITLE>计数器</TITLE>
<META content=zh-cn http-equiv=Content-Language>
<META content="Microsoft FrontPage 6.0" name=GENERATOR>
<META content=FrontPage.Editor.Document name=ProgId>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="images/default.css" rel=stylesheet type=text/css>
<base target="rightframe">
</HEAD>
<BODY bgColor=#FFFFFF topMargin=0 leftMargin=2>
<TABLE border=0 borderColor=#111111 cellSpacing=0 height="60" id=AutoNumber1 
style="BORDER-COLLAPSE: collapse" width=180>
  <TBODY>
  <TR>
    <TD width="100%" height="11">
      <P style="LINE-HEIGHT: 200%" align="center"><FONT color=#0099FF>
		<a target="_blank" title="在线人数统计" href="http://www.yuchonghu.com">
		<span style="text-decoration: none"><font color="#0099FF">当前在线[<%=i%>]</font></span></a> </FONT><FONT color=#ffffff> <BR></FONT>
		<font color="#0099FF"><a target="_blank" href="http://www.yuchonghu.com">
		<font color="#0099FF"><span style="text-decoration: none">您是本站第</span></font></a><a target="_blank" title="总访问人数统计" href="http://www.yuchonghu.com"><span style="text-decoration: none"><font color="#FF0000"> <%=cnt%> 
		</font><font color="#0099FF">位访问者</font></span></a></font></P></TD></TR></TBODY></TABLE></BODY></HTML>