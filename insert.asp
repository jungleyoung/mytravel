<!--#include file="conn.asp"-->


<%
 tmpno=request.form("textno")
 tmpname=request.form("yourname")
 tmppwd=request.form("pwd")
 
 tempsex=request.form("sex")
 tmphobby=request.form("hobby")
 tmpclass=request.form("class")
 'set cn=server.createobject("adodb.connection")
 'cn.open "driver={sql server};server=jcb;database=tms;uid=sa;pwd=123"
 
 set rs=server.createobject("adodb.recordset")
 selstr="select * from text4 where sno='" & tmpno & "'"
 rs.open selstr,cn

 if not (rs.eof) then
   	response.write "<script language=JavaScript>" & "alert('该账号已被占用，请重新选择账号！');" & "history.back();" & "</script>" 
    rs.close
    set rs=nothing
    cn.close
    set cn=nothing
 else
   'set cn=server.createobject("adodb.connection")
   'cn.open "driver={sql server};server=jcb;database=tms;uid=sa;pwd=123"
   
   strin= "insert into text4(sno,syourname,spwd,ssex,shobby,sclass)"
   strin=strin&"values('"&tmpno&"','"&tmpname&"','"&tmppwd&"','"&tempsex&"','"&tmphobby&"','"&tmpclass&"')"
   cn.execute  strin
   response.write "<script language=JavaScript>" & "alert('注册成功！');" & "history.back();" & "</script>" 
 end if
%>
