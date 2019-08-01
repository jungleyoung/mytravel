<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%Response.Charset="utf-8"%>
<HTML>
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8" /></head>
<body>
<%
'连接ACCESS2003数据库1：
'set cn=server.createobject("adodb.connection")
'strconn="driver={microsoft access driver (*.mdb)};dbq="&server.mappath("tms.mdb")
'cn.open strconn

'连接ACCESS2003数据库2：
'set cn=server.createobject("adodb.connection")
'strconn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("tms.mdb")
'cn.open strconn

'连接SQL Server2000数据库1：
'set cn=server.createobject("adodb.connection")
'strconn="driver={sql server};server=jcb;database=tms;uid=sa;pwd=123"
'cn.open  strconn

'连接SQL Server2000数据库2：
'Set cn = Server.CreateObject("ADODB.Connection")
'connstr="Provider=SQLOLEDB;Data Source=(local);Initial Catalog=tms;User ID=sa;Password=123;" 
'cn.Open connstr 

'连接ACCESS2010数据库1：
set cn=server.createobject("adodb.connection")
strconn="driver={microsoft access driver (*.mdb)};dbq="&server.mappath("tms.mdb")
cn.open strconn


'连接ACCESS2010数据库2：
'strconn="provider=microsoft.ACE.oledb.12.0;data source=" & server.MapPath("tms.accdb")
'set cn = Server.Createobject("ADODB.Connection")
'cn.Open strconn

%>
</body></HTML>