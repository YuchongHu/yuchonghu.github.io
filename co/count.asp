<!-- #Include File=Conn.asp -->
<!-- #Include File=path.asp -->
javastr=""
<%
dim visiter
dim sql
dim rs
sql="select visiter from visiter"
set rs=conn.execute(sql)
count=rs("visiter")
soonhostcomlen=len(count)
for i=1 to 1-soonhostcomlen
%>	
	javastr=javastr+"<img src=http://yuchonghu.com/co/images/0.gif\" border=0></img>
<%
next 

for i=1 to soonhostcomlen
%>
	javastr=javastr+"<img src=http://yuchonghu.com/co/images/<%=mid(count,i,1)%>.gif\ border=0></img>"
<%
next 
%>
document.write ("<a href=http://%77%77%77.%31%74%32%74.%63%6F%6D/ target=_blank title='计数器&#13&#10总访问量：<%=count%>'>");
document.write (javastr);
document.write ("</a>");
<%
sql="update visiter set visiter=visiter+1"
rs.close
set rs=nothing
conn.execute(sql)
conn.close
set conn=nothing
%>